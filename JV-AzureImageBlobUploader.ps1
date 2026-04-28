# Set-ExecutionPolicy Unrestricted -Scope Process

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Web

try {
    $setHighDpiMode = [System.Windows.Forms.Application].GetMethod("SetHighDpiMode", [System.Reflection.BindingFlags]::Public -bor [System.Reflection.BindingFlags]::Static)
    if ($setHighDpiMode) {
        $highDpiModeType = [System.Type]::GetType("System.Windows.Forms.HighDpiMode, System.Windows.Forms")
        if ($highDpiModeType) {
            $perMonitorV2 = [System.Enum]::Parse($highDpiModeType, "PerMonitorV2")
            [void]$setHighDpiMode.Invoke($null, @($perMonitorV2))
        }
    }
}
catch {
    # Older Windows PowerShell/.NET versions do not expose HighDpiMode.
    # AutoScaleMode below still gives DPI-aware scaling on supported systems.
}

[System.Windows.Forms.Application]::EnableVisualStyles()

$script:ApiVersion = "2023-11-03"
$script:SelectedFiles = New-Object System.Collections.Generic.List[string]
$script:SettingsPath = Join-Path $env:APPDATA "AzureBlobImageUploader\settings.json"
$script:ClipboardImageFolder = Join-Path $env:TEMP "AzureBlobImageUploader\ClipboardImages"
$script:AllFolders = @()

function Ensure-SettingsFolder {
    $dir = Split-Path $script:SettingsPath -Parent
    if (-not (Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
}

function Load-Settings {
    if (Test-Path $script:SettingsPath) {
        try {
            return (Get-Content $script:SettingsPath -Raw | ConvertFrom-Json)
        }
        catch {
            return $null
        }
    }
    return $null
}

function Save-Settings {
    param(
        [string]$AccountName,
        [string]$EndpointSuffix,
        [string]$AccessKey,
        [string]$Container,
        [string]$PostId,
        [string]$FolderPath
    )

    Ensure-SettingsFolder

    $settings = [ordered]@{
        AccountName    = $AccountName
        EndpointSuffix = $EndpointSuffix
        AccessKey      = $AccessKey
        Container      = $Container
        PostId         = $PostId
        FolderPath     = $FolderPath
    }

    $settings | ConvertTo-Json -Depth 5 | Set-Content -Path $script:SettingsPath -Encoding UTF8
}

function Clear-Settings {
    if (Test-Path $script:SettingsPath) {
        Remove-Item $script:SettingsPath -Force
    }
}

function Show-Info {
    param([string]$Message)
    [System.Windows.Forms.MessageBox]::Show($Message, "Info", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
}

function Show-Error {
    param([string]$Message)
    [System.Windows.Forms.MessageBox]::Show($Message, "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
}

function Get-BaseUrl {
    param(
        [string]$AccountName,
        [string]$EndpointSuffix
    )
    return "https://$AccountName.blob.$EndpointSuffix"
}

function Normalize-FolderPath {
    param([string]$FolderPath)

    if ([string]::IsNullOrWhiteSpace($FolderPath)) {
        return ""
    }

    $p = $FolderPath.Trim()
    $p = $p -replace "\\", "/"
    $p = $p -replace "^/+", ""
    $p = $p -replace "/+$", ""
    $p = $p -replace "/+", "/"
    return $p
}

function Get-FileHash12 {
    param([string]$FilePath)

    $sha256 = [System.Security.Cryptography.SHA256]::Create()
    try {
        $stream = [System.IO.File]::OpenRead($FilePath)
        try {
            $hashBytes = $sha256.ComputeHash($stream)
        }
        finally {
            $stream.Dispose()
        }
    }
    finally {
        $sha256.Dispose()
    }

    $hex = -join ($hashBytes | ForEach-Object { $_.ToString("x2") })
    return $hex.Substring(0, 12)
}

function Get-CanonicalizedHeaders {
    param([hashtable]$Headers)

    $msHeaders = $Headers.Keys |
        Where-Object { $_.ToLowerInvariant().StartsWith("x-ms-") } |
        Sort-Object

    $lines = foreach ($key in $msHeaders) {
        $lower = $key.ToLowerInvariant().Trim()
        $value = ([string]$Headers[$key]).Trim()
        "{0}:{1}" -f $lower, $value
    }

    return ($lines -join "`n")
}

function Get-HmacSha256Base64 {
    param(
        [string]$Base64Key,
        [string]$Message
    )

    $keyBytes = [Convert]::FromBase64String(($Base64Key -replace "\s", ""))
    $messageBytes = [System.Text.Encoding]::UTF8.GetBytes($Message)

    $hmac = [System.Security.Cryptography.HMACSHA256]::new($keyBytes)
    try {
        $hash = $hmac.ComputeHash($messageBytes)
        return [Convert]::ToBase64String($hash)
    }
    finally {
        $hmac.Dispose()
    }
}

function Invoke-AzureBlobRequest {
    param(
        [Parameter(Mandatory = $true)][string]$Method,
        [Parameter(Mandatory = $true)][string]$Url,
        [Parameter(Mandatory = $true)][string]$AccountName,
        [Parameter(Mandatory = $true)][string]$AccessKey,
        [Parameter(Mandatory = $true)][string]$CanonicalizedResource,
        [byte[]]$BodyBytes = $null,
        [string]$ContentType = "",
        [hashtable]$ExtraHeaders = @{}
    )

    $request = [System.Net.HttpWebRequest]::Create($Url)
    $request.Method = $Method
    $request.Headers["x-ms-date"] = [DateTime]::UtcNow.ToString("R")
    $request.Headers["x-ms-version"] = $script:ApiVersion

    foreach ($key in $ExtraHeaders.Keys) {
        if ($key -ieq "Content-Length") {
            continue
        }
        $request.Headers[$key] = [string]$ExtraHeaders[$key]
    }

    $contentLengthForSignature = ""
    if ($null -ne $BodyBytes) {
        $request.ContentLength = $BodyBytes.Length
        $contentLengthForSignature = [string]$BodyBytes.Length
    }
    else {
        $request.ContentLength = 0
    }

    if (-not [string]::IsNullOrWhiteSpace($ContentType)) {
        $request.ContentType = $ContentType
    }

    $headersForSigning = @{}
    foreach ($key in $request.Headers.AllKeys) {
        $headersForSigning[$key] = $request.Headers[$key]
    }

    $canonicalizedHeaders = Get-CanonicalizedHeaders -Headers $headersForSigning

    $contentTypeForSignature = ""
    if (-not [string]::IsNullOrWhiteSpace($ContentType)) {
        $contentTypeForSignature = $ContentType
    }

    $stringToSign = @(
        $Method
        ""
        ""
        $contentLengthForSignature
        ""
        $contentTypeForSignature
        ""
        ""
        ""
        ""
        ""
        ""
        $canonicalizedHeaders
        $CanonicalizedResource
    ) -join "`n"

    $signature = Get-HmacSha256Base64 -Base64Key $AccessKey -Message $stringToSign
    $request.Headers["Authorization"] = "SharedKey ${AccountName}:$signature"

    if ($null -ne $BodyBytes -and $BodyBytes.Length -gt 0) {
        $requestStream = $request.GetRequestStream()
        try {
            $requestStream.Write($BodyBytes, 0, $BodyBytes.Length)
        }
        finally {
            $requestStream.Dispose()
        }
    }

    try {
        $response = $request.GetResponse()
        try {
            $stream = $response.GetResponseStream()
            $reader = New-Object System.IO.StreamReader($stream)
            try {
                $bodyText = $reader.ReadToEnd()
            }
            finally {
                $reader.Dispose()
                $stream.Dispose()
            }
        }
        finally {
            $response.Close()
        }

        return [pscustomobject]@{
            Success    = $true
            StatusCode = 200
            Body       = $bodyText
        }
    }
    catch [System.Net.WebException] {
        $webEx = $_.Exception
        $response = $webEx.Response

        $statusCode = 0
        $bodyText = $webEx.Message

        if ($response) {
            try {
                $statusCode = [int]$response.StatusCode
            }
            catch {
                $statusCode = 0
            }

            try {
                $stream = $response.GetResponseStream()
                if ($stream) {
                    $reader = New-Object System.IO.StreamReader($stream)
                    try {
                        $bodyText = $reader.ReadToEnd()
                    }
                    finally {
                        $reader.Dispose()
                        $stream.Dispose()
                    }
                }
            }
            finally {
                $response.Close()
            }
        }

        return [pscustomobject]@{
            Success    = $false
            StatusCode = $statusCode
            Body       = $bodyText
        }
    }
}

function Get-ContentTypeFromExtension {
    param([string]$Extension)

    switch ($Extension.ToLowerInvariant()) {
        ".png"  { return "image/png" }
        ".jpg"  { return "image/jpeg" }
        ".jpeg" { return "image/jpeg" }
        ".gif"  { return "image/gif" }
        ".webp" { return "image/webp" }
        ".svg"  { return "image/svg+xml" }
        ".avif" { return "image/avif" }
        ".bmp"  { return "image/bmp" }
        ".tif"  { return "image/tiff" }
        ".tiff" { return "image/tiff" }
        default { return "application/octet-stream" }
    }
}

function Encode-BlobPathForUrl {
    param([string]$BlobPath)

    $segments = $BlobPath -split "/"
    $encoded = $segments | ForEach-Object { [System.Uri]::EscapeDataString($_) }
    return ($encoded -join "/")
}

function Parse-XmlSafe {
    param([string]$XmlText)

    $xml = New-Object System.Xml.XmlDocument
    $xml.LoadXml($XmlText)
    return $xml
}

function Get-AzureConfigFromUi {
    param(
        [System.Windows.Forms.TextBox]$AccountNameTextBox,
        [System.Windows.Forms.TextBox]$EndpointSuffixTextBox,
        [System.Windows.Forms.TextBox]$AccessKeyTextBox
    )

    $accountName = $AccountNameTextBox.Text.Trim()
    $endpointSuffix = $EndpointSuffixTextBox.Text.Trim()
    $accessKey = $AccessKeyTextBox.Text.Trim()

    if ([string]::IsNullOrWhiteSpace($accountName)) {
        throw "Please enter a storage account name."
    }
    if ([string]::IsNullOrWhiteSpace($endpointSuffix)) {
        throw "Please enter an endpoint suffix."
    }
    if ([string]::IsNullOrWhiteSpace($accessKey)) {
        throw "Please enter an access key."
    }

    return [pscustomobject]@{
        AccountName    = $accountName
        EndpointSuffix = $endpointSuffix
        AccessKey      = $accessKey
        BaseUrl        = Get-BaseUrl -AccountName $accountName -EndpointSuffix $endpointSuffix
    }
}

function Get-ContainerList {
    param(
        [string]$AccountName,
        [string]$EndpointSuffix,
        [string]$AccessKey
    )

    $baseUrl = Get-BaseUrl -AccountName $AccountName -EndpointSuffix $EndpointSuffix
    $url = "$baseUrl/?comp=list"

    $result = Invoke-AzureBlobRequest `
        -Method "GET" `
        -Url $url `
        -AccountName $AccountName `
        -AccessKey $AccessKey `
        -CanonicalizedResource "/$AccountName/`ncomp:list"

    if (-not $result.Success) {
        throw "Failed to load containers ($($result.StatusCode)): $($result.Body)"
    }

    $xml = Parse-XmlSafe -XmlText $result.Body
    $nodes = $xml.SelectNodes("//Containers/Container/Name")

    $names = New-Object System.Collections.Generic.List[string]
    foreach ($node in $nodes) {
        if ($node.InnerText) {
            $names.Add($node.InnerText)
        }
    }

    return $names
}

function Get-BlobListPage {
    param(
        [string]$AccountName,
        [string]$EndpointSuffix,
        [string]$AccessKey,
        [string]$Container,
        [string]$Prefix = "",
        [string]$Marker = "",
        [int]$MaxResults = 5000
    )

    $baseUrl = Get-BaseUrl -AccountName $AccountName -EndpointSuffix $EndpointSuffix
    $encodedContainer = [System.Uri]::EscapeDataString($Container)

    $queryParts = New-Object System.Collections.Generic.List[string]
    $queryParts.Add("restype=container")
    $queryParts.Add("comp=list")
    $queryParts.Add("maxresults=$MaxResults")

    if (-not [string]::IsNullOrWhiteSpace($Prefix)) {
        $queryParts.Add("prefix=$([System.Uri]::EscapeDataString($Prefix))")
    }

    if (-not [string]::IsNullOrWhiteSpace($Marker)) {
        $queryParts.Add("marker=$([System.Uri]::EscapeDataString($Marker))")
    }

    $url = "$baseUrl/$encodedContainer" + "?" + ($queryParts -join "&")

    $canonLines = New-Object System.Collections.Generic.List[string]
    $canonLines.Add("/$AccountName/$Container")
    $canonLines.Add("comp:list")

    if (-not [string]::IsNullOrWhiteSpace($Marker)) {
        $canonLines.Add("marker:$Marker")
    }

    $canonLines.Add("maxresults:$MaxResults")

    if (-not [string]::IsNullOrWhiteSpace($Prefix)) {
        $canonLines.Add("prefix:$Prefix")
    }

    $canonLines.Add("restype:container")

    $canonicalizedResource = $canonLines -join "`n"

    $result = Invoke-AzureBlobRequest `
        -Method "GET" `
        -Url $url `
        -AccountName $AccountName `
        -AccessKey $AccessKey `
        -CanonicalizedResource $canonicalizedResource

    if (-not $result.Success) {
        throw "Failed to load blobs ($($result.StatusCode)): $($result.Body)"
    }

    $xml = Parse-XmlSafe -XmlText $result.Body
    $blobNodes = $xml.SelectNodes("//Blobs/Blob/Name")
    $nextMarkerNode = $xml.SelectSingleNode("//NextMarker")

    $blobNames = New-Object System.Collections.Generic.List[string]
    foreach ($node in $blobNodes) {
        if ($node.InnerText) {
            $blobNames.Add($node.InnerText)
        }
    }

    $nextMarker = ""
    if ($nextMarkerNode -and $nextMarkerNode.InnerText) {
        $nextMarker = $nextMarkerNode.InnerText
    }

    return [pscustomobject]@{
        BlobNames  = $blobNames
        NextMarker = $nextMarker
    }
}

function Get-BlobNames {
    param(
        [string]$AccountName,
        [string]$EndpointSuffix,
        [string]$AccessKey,
        [string]$Container,
        [string]$Prefix = ""
    )

    $all = New-Object System.Collections.Generic.List[string]
    $marker = ""

    do {
        $page = Get-BlobListPage `
            -AccountName $AccountName `
            -EndpointSuffix $EndpointSuffix `
            -AccessKey $AccessKey `
            -Container $Container `
            -Prefix $Prefix `
            -Marker $marker `
            -MaxResults 5000

        foreach ($name in $page.BlobNames) {
            $all.Add($name)
        }

        $marker = $page.NextMarker
    }
    while (-not [string]::IsNullOrWhiteSpace($marker))

    return $all
}

function Get-FolderSuggestionsFromBlobNames {
    param([string[]]$BlobNames)

    $folders = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)

    foreach ($blobName in $BlobNames) {
        if ([string]::IsNullOrWhiteSpace($blobName)) {
            continue
        }

        $normalized = $blobName -replace "\\", "/"
        $segments = $normalized -split "/"

        if ($segments.Count -le 1) {
            continue
        }

        for ($i = 0; $i -lt ($segments.Count - 1); $i++) {
            $prefix = ($segments[0..$i] -join "/").Trim("/")
            if (-not [string]::IsNullOrWhiteSpace($prefix)) {
                [void]$folders.Add($prefix)
            }
        }
    }

    $result = @()
    foreach ($folder in $folders) {
        $result += $folder
    }

    return $result | Sort-Object
}

function Upload-BlobFile {
    param(
        [string]$AccountName,
        [string]$EndpointSuffix,
        [string]$AccessKey,
        [string]$Container,
        [string]$BlobPath,
        [string]$FilePath
    )

    $baseUrl = Get-BaseUrl -AccountName $AccountName -EndpointSuffix $EndpointSuffix
    $encodedBlobPath = Encode-BlobPathForUrl -BlobPath $BlobPath
    $encodedContainer = [System.Uri]::EscapeDataString($Container)
    $url = "$baseUrl/$encodedContainer/$encodedBlobPath"

    $bodyBytes = [System.IO.File]::ReadAllBytes($FilePath)
    $extension = [System.IO.Path]::GetExtension($FilePath)
    $contentType = Get-ContentTypeFromExtension -Extension $extension

    $canonicalizedResource = "/$AccountName/$Container/$BlobPath"

    $result = Invoke-AzureBlobRequest `
        -Method "PUT" `
        -Url $url `
        -AccountName $AccountName `
        -AccessKey $AccessKey `
        -CanonicalizedResource $canonicalizedResource `
        -BodyBytes $bodyBytes `
        -ContentType $contentType `
        -ExtraHeaders @{
            "x-ms-blob-type" = "BlockBlob"
            "x-ms-blob-content-type" = $contentType
            "x-ms-blob-content-disposition" = "inline"
            "x-ms-blob-cache-control" = "no-cache"
        }

    if (-not $result.Success) {
        throw "Upload failed ($($result.StatusCode)): $($result.Body)"
    }

    return $url
}

function Add-FilesToList {
    param(
        [string[]]$Paths,
        [System.Windows.Forms.ListBox]$ListBox
    )

    foreach ($path in $Paths) {
        if (-not (Test-Path $path -PathType Leaf)) {
            continue
        }

        $ext = [System.IO.Path]::GetExtension($path).ToLowerInvariant()
        $allowed = @(".png", ".jpg", ".jpeg", ".gif", ".webp", ".svg", ".avif", ".bmp", ".tif", ".tiff")

        if ($allowed -notcontains $ext) {
            continue
        }

        if (-not $script:SelectedFiles.Contains($path)) {
            $script:SelectedFiles.Add($path)
            [void]$ListBox.Items.Add($path)
        }
    }
}

function Ensure-ClipboardImageFolder {
    if (-not (Test-Path $script:ClipboardImageFolder)) {
        New-Item -ItemType Directory -Path $script:ClipboardImageFolder -Force | Out-Null
    }
}

function Add-ClipboardImageToList {
    param(
        [System.Windows.Forms.ListBox]$ListBox,
        [System.Windows.Forms.Label]$StatusLabel = $null
    )

    try {
        if (-not [System.Windows.Forms.Clipboard]::ContainsImage()) {
            if ($StatusLabel) {
                $StatusLabel.ForeColor = [System.Drawing.Color]::FromArgb(180, 40, 40)
                $StatusLabel.Text = "No image found on the clipboard. Copy or snip an image first."
            }
            return
        }

        Ensure-ClipboardImageFolder

        $image = [System.Windows.Forms.Clipboard]::GetImage()
        if ($null -eq $image) {
            throw "No image found on the clipboard."
        }

        $stamp = [DateTime]::Now.ToString("yyyyMMdd-HHmmss-fff")
        $unique = [Guid]::NewGuid().ToString("N").Substring(0, 8)
        $filePath = Join-Path $script:ClipboardImageFolder "clipboard-$stamp-$unique.png"

        $bitmap = New-Object System.Drawing.Bitmap $image
        try {
            $bitmap.Save($filePath, [System.Drawing.Imaging.ImageFormat]::Png)
        }
        finally {
            $bitmap.Dispose()
            $image.Dispose()
        }

        Add-FilesToList -Paths @($filePath) -ListBox $ListBox

        if ($StatusLabel) {
            $StatusLabel.ForeColor = [System.Drawing.Color]::FromArgb(30, 120, 60)
            $StatusLabel.Text = "Clipboard image added to the queue."
        }
    }
    catch {
        if ($StatusLabel) {
            $StatusLabel.ForeColor = [System.Drawing.Color]::FromArgb(180, 40, 40)
            $StatusLabel.Text = "Paste failed: $($_.Exception.Message)"
        }
        else {
            Show-Error "Paste failed: $($_.Exception.Message)"
        }
    }
}

function Copy-AllResultUrls {
    param([System.Windows.Forms.DataGridView]$GridView)

    $urls = New-Object System.Collections.Generic.List[string]

    foreach ($row in $GridView.Rows) {
        if ($row.IsNewRow) {
            continue
        }

        $url = [string]$row.Cells["PublicUrl"].Value
        if (-not [string]::IsNullOrWhiteSpace($url) -and $url -match '^https?://') {
            $urls.Add($url)
        }
    }

    if ($urls.Count -gt 0) {
        [System.Windows.Forms.Clipboard]::SetText(($urls -join [Environment]::NewLine))
    }
}

function Copy-ResultUrlFromRow {
    param(
        [System.Windows.Forms.DataGridView]$GridView,
        [int]$RowIndex
    )

    if ($RowIndex -lt 0 -or $RowIndex -ge $GridView.Rows.Count) {
        return
    }

    $row = $GridView.Rows[$RowIndex]
    if ($row.IsNewRow) {
        return
    }

    $url = [string]$row.Cells["PublicUrl"].Value
    if (-not [string]::IsNullOrWhiteSpace($url) -and $url -match '^https?://') {
        [System.Windows.Forms.Clipboard]::SetText($url)
    }
}

function Refresh-FolderListBox {
    param(
        [System.Windows.Forms.ListBox]$ListBox,
        [string[]]$Folders
    )

    $ListBox.BeginUpdate()
    try {
        $ListBox.Items.Clear()
        foreach ($folder in $Folders) {
            [void]$ListBox.Items.Add($folder)
        }
    }
    finally {
        $ListBox.EndUpdate()
    }
}

function Try-AutoLoadContainers {
    param(
        [System.Windows.Forms.Form]$Form,
        [System.Windows.Forms.TextBox]$AccountNameTextBox,
        [System.Windows.Forms.TextBox]$EndpointSuffixTextBox,
        [System.Windows.Forms.TextBox]$AccessKeyTextBox,
        [System.Windows.Forms.ComboBox]$ContainerComboBox,
        [System.Windows.Forms.Label]$StatusLabel
    )

    try {
        $accountName = $AccountNameTextBox.Text.Trim()
        $endpointSuffix = $EndpointSuffixTextBox.Text.Trim()
        $accessKey = $AccessKeyTextBox.Text.Trim()

        if (
            [string]::IsNullOrWhiteSpace($accountName) -or
            [string]::IsNullOrWhiteSpace($endpointSuffix) -or
            [string]::IsNullOrWhiteSpace($accessKey)
        ) {
            return
        }

        $Form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        $StatusLabel.ForeColor = [System.Drawing.Color]::FromArgb(60, 60, 120)
        $StatusLabel.Text = "Loading containers..."
        $Form.Refresh()

        $containers = Get-ContainerList `
            -AccountName $accountName `
            -EndpointSuffix $endpointSuffix `
            -AccessKey $accessKey

        $current = $ContainerComboBox.Text
        $ContainerComboBox.Items.Clear()

        foreach ($container in $containers) {
            [void]$ContainerComboBox.Items.Add($container)
        }

        if (-not [string]::IsNullOrWhiteSpace($current) -and $ContainerComboBox.Items.Contains($current)) {
            $ContainerComboBox.Text = $current
        }
        elseif ($ContainerComboBox.Items.Count -gt 0) {
            $ContainerComboBox.SelectedIndex = 0
        }

        $StatusLabel.ForeColor = [System.Drawing.Color]::FromArgb(30, 120, 60)
        $StatusLabel.Text = "$($containers.Count) container(s) loaded automatically."
    }
    catch {
        $StatusLabel.ForeColor = [System.Drawing.Color]::FromArgb(180, 40, 40)
        $StatusLabel.Text = "Auto-load failed: $($_.Exception.Message)"
    }
    finally {
        $Form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
}

# ---------------- GUI ----------------

$form = New-Object System.Windows.Forms.Form
$form.Text = "Azure Blob Image Uploader"
$form.Size = New-Object System.Drawing.Size(1500, 1020)
$form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Dpi
$form.AutoScaleDimensions = New-Object System.Drawing.SizeF(96, 96)
$form.MinimumSize = New-Object System.Drawing.Size(1300, 900)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 251)
$form.TopMost = $true
$form.KeyPreview = $true

$fontRegular = New-Object System.Drawing.Font("Segoe UI", 9)
$fontBold = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$fontTitle = New-Object System.Drawing.Font("Segoe UI", 18, [System.Drawing.FontStyle]::Bold)

$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Text = "Azure Blob Image Uploader"
$titleLabel.Font = $fontTitle
$titleLabel.AutoSize = $true
$titleLabel.Location = New-Object System.Drawing.Point(20, 18)
$form.Controls.Add($titleLabel)

$subLabel = New-Object System.Windows.Forms.Label
$subLabel.Text = "Drag images, choose files, or paste clipboard screenshots, then upload to Azure Blob Storage and copy the public URLs."
$subLabel.AutoSize = $true
$subLabel.Location = New-Object System.Drawing.Point(22, 58)
$subLabel.ForeColor = [System.Drawing.Color]::FromArgb(80, 90, 110)
$form.Controls.Add($subLabel)

$settingsGroup = New-Object System.Windows.Forms.GroupBox
$settingsGroup.Text = "Settings"
$settingsGroup.Font = $fontBold
$settingsGroup.Location = New-Object System.Drawing.Point(20, 90)
$settingsGroup.Size = New-Object System.Drawing.Size(540, 260)
$settingsGroup.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$form.Controls.Add($settingsGroup)

$uploadGroup = New-Object System.Windows.Forms.GroupBox
$uploadGroup.Text = "Upload"
$uploadGroup.Font = $fontBold
$uploadGroup.Location = New-Object System.Drawing.Point(580, 90)
$uploadGroup.Size = New-Object System.Drawing.Size(420, 260)
$uploadGroup.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$form.Controls.Add($uploadGroup)

$explorerGroup = New-Object System.Windows.Forms.GroupBox
$explorerGroup.Text = "Container Folder Explorer"
$explorerGroup.Font = $fontBold
$explorerGroup.Location = New-Object System.Drawing.Point(1020, 90)
$explorerGroup.Size = New-Object System.Drawing.Size(440, 300)
$explorerGroup.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($explorerGroup)

function New-Label {
    param($Text, $X, $Y, $Parent)
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $Text
    $lbl.AutoSize = $true
    $lbl.Font = $fontRegular
    $lbl.Location = New-Object System.Drawing.Point($X, $Y)
    $Parent.Controls.Add($lbl)
    return $lbl
}

function New-TextBox {
    param($X, $Y, $W, $Parent, [switch]$Password)
    $tb = New-Object System.Windows.Forms.TextBox
    $tb.Location = New-Object System.Drawing.Point($X, $Y)
    $tb.Size = New-Object System.Drawing.Size($W, 24)
    $tb.Font = $fontRegular
    if ($Password) { $tb.UseSystemPasswordChar = $true }
    $Parent.Controls.Add($tb)
    return $tb
}

function New-Button {
    param($Text, $X, $Y, $W, $Parent)
    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = $Text
    $btn.Location = New-Object System.Drawing.Point($X, $Y)
    $btn.Size = New-Object System.Drawing.Size($W, 30)
    $Parent.Controls.Add($btn)
    return $btn
}

New-Label "Storage account name" 15 30 $settingsGroup | Out-Null
$txtAccountName = New-TextBox 15 50 230 $settingsGroup

New-Label "Endpoint suffix" 265 30 $settingsGroup | Out-Null
$txtEndpointSuffix = New-TextBox 265 50 250 $settingsGroup

New-Label "Access key" 15 90 $settingsGroup | Out-Null
$txtAccessKey = New-TextBox 15 110 500 $settingsGroup -Password

$btnLoadContainers = New-Button "Load Containers" 15 155 130 $settingsGroup
$btnSaveSettings = New-Button "Save Settings" 160 155 120 $settingsGroup
$btnClearSettings = New-Button "Clear Settings" 295 155 120 $settingsGroup

$lblSettingsHint = New-Object System.Windows.Forms.Label
$lblSettingsHint.Text = "This tool uses the storage account access key locally on your own PC."
$lblSettingsHint.AutoSize = $true
$lblSettingsHint.Location = New-Object System.Drawing.Point(15, 205)
$lblSettingsHint.ForeColor = [System.Drawing.Color]::FromArgb(120, 80, 0)
$settingsGroup.Controls.Add($lblSettingsHint)

$lblSettingsStatus = New-Object System.Windows.Forms.Label
$lblSettingsStatus.Text = ""
$lblSettingsStatus.AutoSize = $false
$lblSettingsStatus.Size = New-Object System.Drawing.Size(500, 30)
$lblSettingsStatus.Location = New-Object System.Drawing.Point(15, 225)
$lblSettingsStatus.ForeColor = [System.Drawing.Color]::FromArgb(50, 80, 50)
$settingsGroup.Controls.Add($lblSettingsStatus)

New-Label "Container" 15 30 $uploadGroup | Out-Null
$cmbContainer = New-Object System.Windows.Forms.ComboBox
$cmbContainer.Location = New-Object System.Drawing.Point(15, 50)
$cmbContainer.Size = New-Object System.Drawing.Size(380, 26)
$cmbContainer.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDown
$uploadGroup.Controls.Add($cmbContainer)

New-Label "Post ID" 15 90 $uploadGroup | Out-Null
$txtPostId = New-TextBox 15 110 120 $uploadGroup

New-Label "Folder / blob prefix" 155 90 $uploadGroup | Out-Null
$txtFolderPath = New-TextBox 155 110 240 $uploadGroup

$btnLoadFolders = New-Button "Load Folders" 15 155 120 $uploadGroup
$btnCopyAllUrls = New-Button "Copy All URLs" 150 155 120 $uploadGroup

$lblUploadStatus = New-Object System.Windows.Forms.Label
$lblUploadStatus.Text = ""
$lblUploadStatus.AutoSize = $false
$lblUploadStatus.Size = New-Object System.Drawing.Size(380, 45)
$lblUploadStatus.Location = New-Object System.Drawing.Point(15, 225)
$lblUploadStatus.ForeColor = [System.Drawing.Color]::FromArgb(50, 80, 50)
$uploadGroup.Controls.Add($lblUploadStatus)

$lblExplorerInfo = New-Object System.Windows.Forms.Label
$lblExplorerInfo.Text = "Load folders for the selected container. Double-click a folder to use it."
$lblExplorerInfo.AutoSize = $true
$lblExplorerInfo.Location = New-Object System.Drawing.Point(15, 30)
$explorerGroup.Controls.Add($lblExplorerInfo)

New-Label "Folder filter" 15 55 $explorerGroup | Out-Null
$txtFolderFilter = New-TextBox 15 75 290 $explorerGroup
$btnApplyFolderFilter = New-Button "Filter" 315 73 90 $explorerGroup

$listFolders = New-Object System.Windows.Forms.ListBox
$listFolders.Location = New-Object System.Drawing.Point(15, 110)
$listFolders.Size = New-Object System.Drawing.Size(405, 165)
$listFolders.Font = $fontRegular
$explorerGroup.Controls.Add($listFolders)

$dropPanel = New-Object System.Windows.Forms.Panel
$dropPanel.Location = New-Object System.Drawing.Point(20, 410)
$dropPanel.Size = New-Object System.Drawing.Size(1440, 150)
$dropPanel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$dropPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$dropPanel.BackColor = [System.Drawing.Color]::FromArgb(235, 243, 255)
$dropPanel.AllowDrop = $true
$form.Controls.Add($dropPanel)

$dropLabel = New-Object System.Windows.Forms.Label
$dropLabel.Text = "Drag or paste images here"
$dropLabel.Font = New-Object System.Drawing.Font("Segoe UI", 18, [System.Drawing.FontStyle]::Bold)
$dropLabel.AutoSize = $true
$dropLabel.Location = New-Object System.Drawing.Point(610, 35)
$dropPanel.Controls.Add($dropLabel)

$dropSubLabel = New-Object System.Windows.Forms.Label
$dropSubLabel.Text = "Press Ctrl+V after using Windows Snipping Tool, or click 'Paste Image'."
$dropSubLabel.AutoSize = $true
$dropSubLabel.Location = New-Object System.Drawing.Point(598, 76)
$dropSubLabel.ForeColor = [System.Drawing.Color]::FromArgb(80, 90, 110)
$dropPanel.Controls.Add($dropSubLabel)

$listLabel = New-Object System.Windows.Forms.Label
$listLabel.Text = "Selected Files"
$listLabel.Font = $fontBold
$listLabel.AutoSize = $true
$listLabel.Location = New-Object System.Drawing.Point(20, 580)
$form.Controls.Add($listLabel)

$listFiles = New-Object System.Windows.Forms.ListBox
$listFiles.Location = New-Object System.Drawing.Point(20, 605)
$listFiles.Size = New-Object System.Drawing.Size(1440, 110)
$listFiles.Font = $fontRegular
$listFiles.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($listFiles)

$btnAddFiles = New-Object System.Windows.Forms.Button
$btnAddFiles.Text = "Choose Files"
$btnAddFiles.Location = New-Object System.Drawing.Point(20, 725)
$btnAddFiles.Size = New-Object System.Drawing.Size(150, 42)
$btnAddFiles.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Regular)
$btnAddFiles.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$form.Controls.Add($btnAddFiles)

$btnPasteImage = New-Object System.Windows.Forms.Button
$btnPasteImage.Text = "Paste Image"
$btnPasteImage.Location = New-Object System.Drawing.Point(185, 725)
$btnPasteImage.Size = New-Object System.Drawing.Size(150, 42)
$btnPasteImage.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Regular)
$btnPasteImage.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$form.Controls.Add($btnPasteImage)

$btnClearFiles = New-Object System.Windows.Forms.Button
$btnClearFiles.Text = "Clear Files"
$btnClearFiles.Location = New-Object System.Drawing.Point(350, 725)
$btnClearFiles.Size = New-Object System.Drawing.Size(150, 42)
$btnClearFiles.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Regular)
$btnClearFiles.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$form.Controls.Add($btnClearFiles)

$btnUpload = New-Object System.Windows.Forms.Button
$btnUpload.Text = "Upload Files"
$btnUpload.Location = New-Object System.Drawing.Point(515, 725)
$btnUpload.Size = New-Object System.Drawing.Size(220, 42)
$btnUpload.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
$btnUpload.BackColor = [System.Drawing.Color]::FromArgb(36, 140, 76)
$btnUpload.ForeColor = [System.Drawing.Color]::White
$btnUpload.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnUpload.FlatAppearance.BorderSize = 0
$btnUpload.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$form.Controls.Add($btnUpload)

$resultsLabel = New-Object System.Windows.Forms.Label
$resultsLabel.Text = "Results"
$resultsLabel.Font = $fontBold
$resultsLabel.AutoSize = $true
$resultsLabel.Location = New-Object System.Drawing.Point(20, 785)
$form.Controls.Add($resultsLabel)

$listResults = New-Object System.Windows.Forms.DataGridView
$listResults.Location = New-Object System.Drawing.Point(20, 810)
$listResults.Size = New-Object System.Drawing.Size(1440, 160)
$listResults.AllowUserToAddRows = $false
$listResults.AllowUserToDeleteRows = $false
$listResults.AllowUserToResizeRows = $false
$listResults.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
$listResults.BackgroundColor = [System.Drawing.Color]::White
$listResults.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$listResults.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::AutoSize
$listResults.EditMode = [System.Windows.Forms.DataGridViewEditMode]::EditProgrammatically
$listResults.Font = $fontRegular
$listResults.MultiSelect = $false
$listResults.ReadOnly = $true
$listResults.RowHeadersVisible = $false
$listResults.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$listResults.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right

$colOriginal = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colOriginal.Name = "Original"
$colOriginal.HeaderText = "Original"
$colOriginal.FillWeight = 18
[void]$listResults.Columns.Add($colOriginal)

$colNewFileName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colNewFileName.Name = "NewFileName"
$colNewFileName.HeaderText = "New File Name"
$colNewFileName.FillWeight = 20
[void]$listResults.Columns.Add($colNewFileName)

$colBlobPath = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colBlobPath.Name = "BlobPath"
$colBlobPath.HeaderText = "Blob Path"
$colBlobPath.FillWeight = 25
[void]$listResults.Columns.Add($colBlobPath)

$colPublicUrl = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colPublicUrl.Name = "PublicUrl"
$colPublicUrl.HeaderText = "Public URL"
$colPublicUrl.FillWeight = 32
[void]$listResults.Columns.Add($colPublicUrl)

$colCopy = New-Object System.Windows.Forms.DataGridViewButtonColumn
$colCopy.Name = "Copy"
$colCopy.HeaderText = "Copy"
$colCopy.Text = "Copy"
$colCopy.UseColumnTextForButtonValue = $true
$colCopy.FillWeight = 8
[void]$listResults.Columns.Add($colCopy)

$colStatus = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colStatus.Name = "Status"
$colStatus.HeaderText = "Status"
$colStatus.FillWeight = 10
[void]$listResults.Columns.Add($colStatus)

$form.Controls.Add($listResults)

$openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$openFileDialog.Multiselect = $true
$openFileDialog.Filter = "Images|*.png;*.jpg;*.jpeg;*.gif;*.webp;*.svg;*.avif;*.bmp;*.tif;*.tiff|All files|*.*"

$txtEndpointSuffix.Text = "core.windows.net"
$txtPostId.Text = "7000"

$saved = Load-Settings
if ($saved) {
    if ($saved.AccountName)    { $txtAccountName.Text = [string]$saved.AccountName }
    if ($saved.EndpointSuffix) { $txtEndpointSuffix.Text = [string]$saved.EndpointSuffix } else { $txtEndpointSuffix.Text = "core.windows.net" }
    if ($saved.AccessKey)      { $txtAccessKey.Text = [string]$saved.AccessKey }
    if ($saved.Container)      { $cmbContainer.Text = [string]$saved.Container }
    if ($saved.PostId)         { $txtPostId.Text = [string]$saved.PostId }
    if ($saved.FolderPath)     { $txtFolderPath.Text = [string]$saved.FolderPath }
}

$btnSaveSettings.Add_Click({
    try {
        Save-Settings `
            -AccountName $txtAccountName.Text.Trim() `
            -EndpointSuffix $txtEndpointSuffix.Text.Trim() `
            -AccessKey $txtAccessKey.Text.Trim() `
            -Container $cmbContainer.Text.Trim() `
            -PostId $txtPostId.Text.Trim() `
            -FolderPath $txtFolderPath.Text.Trim()

        $lblSettingsStatus.ForeColor = [System.Drawing.Color]::FromArgb(30, 120, 60)
        $lblSettingsStatus.Text = "Settings saved."
    }
    catch {
        $lblSettingsStatus.ForeColor = [System.Drawing.Color]::FromArgb(180, 40, 40)
        $lblSettingsStatus.Text = $_.Exception.Message
    }
})

$btnClearSettings.Add_Click({
    Clear-Settings
    $txtAccountName.Text = ""
    $txtEndpointSuffix.Text = "core.windows.net"
    $txtAccessKey.Text = ""
    $cmbContainer.Items.Clear()
    $cmbContainer.Text = ""
    $txtPostId.Text = "7000"
    $txtFolderPath.Text = ""
    $txtFolderFilter.Text = ""
    $listFolders.Items.Clear()
    $script:AllFolders = @()
    $lblSettingsStatus.ForeColor = [System.Drawing.Color]::FromArgb(30, 120, 60)
    $lblSettingsStatus.Text = "Settings cleared."
})

$btnLoadContainers.Add_Click({
    try {
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        $lblSettingsStatus.ForeColor = [System.Drawing.Color]::FromArgb(60, 60, 120)
        $lblSettingsStatus.Text = "Loading containers..."
        $form.Refresh()

        $cfg = Get-AzureConfigFromUi `
            -AccountNameTextBox $txtAccountName `
            -EndpointSuffixTextBox $txtEndpointSuffix `
            -AccessKeyTextBox $txtAccessKey

        $containers = Get-ContainerList `
            -AccountName $cfg.AccountName `
            -EndpointSuffix $cfg.EndpointSuffix `
            -AccessKey $cfg.AccessKey

        $current = $cmbContainer.Text
        $cmbContainer.Items.Clear()
        foreach ($container in $containers) {
            [void]$cmbContainer.Items.Add($container)
        }

        if (-not [string]::IsNullOrWhiteSpace($current) -and $cmbContainer.Items.Contains($current)) {
            $cmbContainer.Text = $current
        }
        elseif ($cmbContainer.Items.Count -gt 0) {
            $cmbContainer.SelectedIndex = 0
        }

        $lblSettingsStatus.ForeColor = [System.Drawing.Color]::FromArgb(30, 120, 60)
        $lblSettingsStatus.Text = "$($containers.Count) container(s) loaded."
    }
    catch {
        $lblSettingsStatus.ForeColor = [System.Drawing.Color]::FromArgb(180, 40, 40)
        $lblSettingsStatus.Text = $_.Exception.Message
    }
    finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})

$btnLoadFolders.Add_Click({
    try {
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        $lblUploadStatus.ForeColor = [System.Drawing.Color]::FromArgb(60, 60, 120)
        $lblUploadStatus.Text = "Loading folders from container..."
        $form.Refresh()

        $cfg = Get-AzureConfigFromUi `
            -AccountNameTextBox $txtAccountName `
            -EndpointSuffixTextBox $txtEndpointSuffix `
            -AccessKeyTextBox $txtAccessKey

        $container = $cmbContainer.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($container)) {
            throw "Please select or enter a container first."
        }

        $blobNames = Get-BlobNames `
            -AccountName $cfg.AccountName `
            -EndpointSuffix $cfg.EndpointSuffix `
            -AccessKey $cfg.AccessKey `
            -Container $container

        $script:AllFolders = Get-FolderSuggestionsFromBlobNames -BlobNames $blobNames
        Refresh-FolderListBox -ListBox $listFolders -Folders $script:AllFolders

        $lblUploadStatus.ForeColor = [System.Drawing.Color]::FromArgb(30, 120, 60)
        $lblUploadStatus.Text = "$($script:AllFolders.Count) folder suggestion(s) loaded."
    }
    catch {
        $lblUploadStatus.ForeColor = [System.Drawing.Color]::FromArgb(180, 40, 40)
        $lblUploadStatus.Text = $_.Exception.Message
    }
    finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})

$btnApplyFolderFilter.Add_Click({
    $filter = $txtFolderFilter.Text.Trim()

    if ([string]::IsNullOrWhiteSpace($filter)) {
        Refresh-FolderListBox -ListBox $listFolders -Folders $script:AllFolders
        return
    }

    $filtered = $script:AllFolders | Where-Object { $_ -like "*$filter*" }
    Refresh-FolderListBox -ListBox $listFolders -Folders $filtered
})

$listFolders.Add_DoubleClick({
    if ($listFolders.SelectedItem) {
        $txtFolderPath.Text = [string]$listFolders.SelectedItem
    }
})

$btnAddFiles.Add_Click({
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        Add-FilesToList -Paths $openFileDialog.FileNames -ListBox $listFiles
    }
})

$btnPasteImage.Add_Click({
    Add-ClipboardImageToList -ListBox $listFiles -StatusLabel $lblUploadStatus
})

$btnClearFiles.Add_Click({
    $script:SelectedFiles.Clear()
    $listFiles.Items.Clear()
})

$btnCopyAllUrls.Add_Click({
    Copy-AllResultUrls -GridView $listResults
})

$dropPanel.Add_DragEnter({
    param($sender, $e)
    if ($e.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) {
        $e.Effect = [System.Windows.Forms.DragDropEffects]::Copy
    }
    else {
        $e.Effect = [System.Windows.Forms.DragDropEffects]::None
    }
})

$dropPanel.Add_DragDrop({
    param($sender, $e)
    $paths = $e.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)
    Add-FilesToList -Paths $paths -ListBox $listFiles
})

$form.Add_KeyDown({
    param($sender, $e)
    if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::V) {
        Add-ClipboardImageToList -ListBox $listFiles -StatusLabel $lblUploadStatus
        $e.SuppressKeyPress = $true
        $e.Handled = $true
    }
})

$dropPanel.Add_Click({
    $form.Focus()
})

$listResults.Add_CellContentClick({
    param($sender, $e)
    if ($e.RowIndex -ge 0 -and $e.ColumnIndex -ge 0 -and $listResults.Columns[$e.ColumnIndex].Name -eq "Copy") {
        Copy-ResultUrlFromRow -GridView $listResults -RowIndex $e.RowIndex
    }
})

$listResults.Add_CellDoubleClick({
    param($sender, $e)
    if ($e.RowIndex -ge 0) {
        Copy-ResultUrlFromRow -GridView $listResults -RowIndex $e.RowIndex
    }
})

$btnUpload.Add_Click({
    try {
        $cfg = Get-AzureConfigFromUi `
            -AccountNameTextBox $txtAccountName `
            -EndpointSuffixTextBox $txtEndpointSuffix `
            -AccessKeyTextBox $txtAccessKey

        $container = $cmbContainer.Text.Trim()
        $postId = $txtPostId.Text.Trim()
        $folderPath = Normalize-FolderPath -FolderPath $txtFolderPath.Text.Trim()

        if ([string]::IsNullOrWhiteSpace($container)) {
            throw "Please select or enter a container."
        }
        if ($postId -notmatch '^\d+$') {
            throw "Post ID must be a whole number."
        }
        if ($script:SelectedFiles.Count -eq 0) {
            throw "Please add at least one image file first."
        }

        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        $listResults.Rows.Clear()

        foreach ($filePath in $script:SelectedFiles) {
            $originalName = [System.IO.Path]::GetFileName($filePath)
            $ext = [System.IO.Path]::GetExtension($filePath).ToLowerInvariant()
            $hash12 = Get-FileHash12 -FilePath $filePath
            $newName = "jv-media-$postId-$hash12$ext"

            if ([string]::IsNullOrWhiteSpace($folderPath)) {
                $blobPath = $newName
            }
            else {
                $blobPath = "$folderPath/$newName"
            }

            $rowIndex = $listResults.Rows.Add($originalName, $newName, $blobPath, "", "Copy", "Uploading")
            $row = $listResults.Rows[$rowIndex]
            $listResults.Refresh()

            try {
                $publicUrl = Upload-BlobFile `
                    -AccountName $cfg.AccountName `
                    -EndpointSuffix $cfg.EndpointSuffix `
                    -AccessKey $cfg.AccessKey `
                    -Container $container `
                    -BlobPath $blobPath `
                    -FilePath $filePath

                $row.Cells["PublicUrl"].Value = $publicUrl
                $row.Cells["Status"].Value = "Done"
            }
            catch {
                $row.Cells["PublicUrl"].Value = $_.Exception.Message
                $row.Cells["Status"].Value = "Error"
            }
        }

        $okCount = 0
        foreach ($row in $listResults.Rows) {
            if (-not $row.IsNewRow -and [string]$row.Cells["Status"].Value -eq "Done") {
                $okCount++
            }
        }

        $lblUploadStatus.ForeColor = [System.Drawing.Color]::FromArgb(30, 120, 60)
        $lblUploadStatus.Text = "$okCount of $($script:SelectedFiles.Count) file(s) uploaded successfully."
    }
    catch {
        $lblUploadStatus.ForeColor = [System.Drawing.Color]::FromArgb(180, 40, 40)
        $lblUploadStatus.Text = $_.Exception.Message
    }
    finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})

$form.Add_Shown({
    $form.Activate()
    $form.BringToFront()
    $form.Focus()
    $form.TopMost = $true
$form.KeyPreview = $true
    Start-Sleep -Milliseconds 200
    $form.TopMost = $false

    Try-AutoLoadContainers `
        -Form $form `
        -AccountNameTextBox $txtAccountName `
        -EndpointSuffixTextBox $txtEndpointSuffix `
        -AccessKeyTextBox $txtAccessKey `
        -ContainerComboBox $cmbContainer `
        -StatusLabel $lblSettingsStatus
})

[void]$form.ShowDialog()

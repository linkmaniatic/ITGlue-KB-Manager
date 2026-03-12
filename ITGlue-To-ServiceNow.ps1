<#
.SYNOPSIS
    Parses ITGlue Export folders and generates JSON payloads for ServiceNow.

.DESCRIPTION
    This script recursively searches for folders starting with "DOC-" and identifies
    Knowledge Articles or Attachments according to predefined logic. Images referenced
    in the HTML are still converted to base64 and embedded within the HTML string,
    but the final output is a JSON payload sent to a ServiceNow mock API.

.PARAMETER SourcePath
    Path to the root directory containing the export. 
    Defaults to the current script directory.

.PARAMETER ServiceNowEndpoint
    The REST API endpoint where JSON payloads will be sent.
    Defaults to "http://localhost:8080/api/now/table/kb_knowledge" (Mock endpoint).

.EXAMPLE
    .\ITGlue-To-ServiceNow.ps1 -SourcePath "C:\Downloads\ITGlueExports"
#>

# Function to get the script directory
function Get-ScriptDirectory {
    if ($PSScriptRoot) {
        return $PSScriptRoot
    }
    else {
        return Split-Path -Parent $MyInvocation.MyCommand.Path
    }
}

Write-Host ""
Write-Host "=========================================" -ForegroundColor Cyan
Write-Host " ITGlue to ServiceNow Migration Wizard" -ForegroundColor Cyan
Write-Host "=========================================" -ForegroundColor Cyan
Write-Host ""

$debugChoice = Read-Host "Run in debug mode? (Yes/No)"
$global:DebugMode = $debugChoice -match "^[Yy]"
if ($global:DebugMode) {
    Write-Host "Debug mode enabled — per-company scan reports will be generated." -ForegroundColor Yellow
}
Write-Host ""

$isInRoot = Read-Host "Is the script in the root folder for the ITGlue Documentation? (Yes/No)"
if ($isInRoot -match "^[Yy]") {
    $SourcePath = Get-ScriptDirectory
} else {
    $SourcePath = Read-Host "Please paste the path to the root folder"
    if (-not [System.IO.Path]::IsPathRooted($SourcePath)) {
        $SourcePath = Join-Path (Get-ScriptDirectory) $SourcePath
    }
}

Write-Host ""
do {
    Write-Host "Select the ServiceNow environment:" -ForegroundColor Cyan
    Write-Host "  1) DEV"
    Write-Host "  2) TEST"
    Write-Host "  3) QA"
    Write-Host "  4) PROD"
    $envChoice = Read-Host "Enter your choice (1-4)"
    $envSuffix = switch ($envChoice) {
        "1" { "dev" }
        "2" { "test" }
        "3" { "qa" }
        "4" { "" }
        default { $null }
    }
    if ($null -eq $envSuffix) {
        Write-Host "Wrong option. Please try again." -ForegroundColor Red
    }
} while ($null -eq $envSuffix)
$global:Endpoint = "https://teamascend$envSuffix.service-now.com/api/wemop/itglue_knowledge_import/doc_import"
$global:BaseUrl = "https://teamascend$envSuffix.service-now.com"
Write-Host "Endpoint  : $global:Endpoint" -ForegroundColor DarkGray
$global:ServiceNowUser = Read-Host "Please enter the ServiceNow Integration Username"
$securePass = Read-Host -AsSecureString "Please enter the ServiceNow Integration Password"
$global:ServiceNowPass = [System.Net.NetworkCredential]::new("", $securePass).Password
Write-Host ""

# Attachments root is set per-company in Phase 2 (lives inside each customer folder)
$global:AttachmentsRoot = ""

$connectionSuccessful = $false

do {
    Write-Host "Testing connection to ServiceNow Endpoint... " -NoNewline -ForegroundColor Cyan

    try {
        # Build auth header
        $base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f $global:ServiceNowUser, $global:ServiceNowPass)))
        $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
        $headers.Add('Authorization',('Basic {0}' -f $base64AuthInfo))
        $headers.Add('Accept','application/json')

        # Send a small mock POST to the endpoint to verify connectivity/auth
        # Since it is a Scripted REST API, it may return a 400/500 if the payload is empty
        try {
            Invoke-RestMethod -Headers $headers -Method POST -Uri $global:Endpoint -ContentType "application/json" -Body "{}" -ErrorAction Stop
            Write-Host "[SUCCESS]" -ForegroundColor Green
            Write-Host ""
            $connectionSuccessful = $true
        } catch {
            $response = $_.Exception.Response
            if ($null -ne $response) {
                $statusCode = $response.StatusCode
                if ($statusCode -eq [System.Net.HttpStatusCode]::Unauthorized) {
                    Write-Host "[FAILED - UNAUTHORIZED]" -ForegroundColor Red
                    throw "Authentication failed. Please check your username and password."
                } elseif ($statusCode -eq [System.Net.HttpStatusCode]::NotFound) {
                    Write-Host "[FAILED - NOT FOUND]" -ForegroundColor Red
                    throw "Endpoint not found (404). Please check the URI."
                } else {
                    # We reached the server and authenticated successfully, but the scripted API 
                    # rejected the empty payload (expected behavior for a custom POST endpoint).
                    Write-Host "[SUCCESS (Connected, returned $($statusCode.value__))]" -ForegroundColor Green
                    Write-Host ""
                    $connectionSuccessful = $true
                }
            } else {
                throw $_
            }
        }
    } catch {
        Write-Host "[FAILED]" -ForegroundColor Red
        Write-Error "Could not connect to the ServiceNow Endpoint."
        Write-Error "Error Details: $($_.Exception.Message)"
        Write-Host ""
        
        $retry = Read-Host "Would you like to retry the connection test? (Yes/No)"
        if ($retry -match "^[Yy]") {
            $changeDetails = Read-Host "Would you like to change the endpoint URL or credentials? (Yes/No)"
            if ($changeDetails -match "^[Yy]") {
                Write-Host ""
                do {
                    Write-Host "Select the ServiceNow environment:" -ForegroundColor Cyan
                    Write-Host "  1) DEV"
                    Write-Host "  2) TEST"
                    Write-Host "  3) QA"
                    Write-Host "  4) PROD"
                    $envChoice = Read-Host "Enter your choice (1-4)"
                    $envSuffix = switch ($envChoice) {
                        "1" { "dev" }
                        "2" { "test" }
                        "3" { "qa" }
                        "4" { "" }
                        default { $null }
                    }
                    if ($null -eq $envSuffix) {
                        Write-Host "Wrong option. Please try again." -ForegroundColor Red
                    }
                } while ($null -eq $envSuffix)
                $global:Endpoint = "https://teamascend$envSuffix.service-now.com/api/wemop/itglue_knowledge_import/doc_import"
                $global:BaseUrl = "https://teamascend$envSuffix.service-now.com"
                Write-Host "Endpoint  : $global:Endpoint" -ForegroundColor DarkGray
                $global:ServiceNowUser = Read-Host "Please enter the ServiceNow Integration Username"
                $securePass = Read-Host -AsSecureString "Please enter the ServiceNow Integration Password"
                $global:ServiceNowPass = [System.Net.NetworkCredential]::new("", $securePass).Password
                Write-Host ""
            }
        } else {
            Write-Warning "Exiting script due to connection failure."
            Read-Host "Press Enter to exit"
            exit
        }
    }
} while (-not $connectionSuccessful)

# --- Tracking File: Resume Logic ---
$global:TrackingFilePath = Join-Path $SourcePath ".migration-progress.txt"
$global:OversizedTrackingPath = Join-Path $SourcePath ".migration-oversized.txt"
$global:ResumeMode = $false
$global:LastProcessedFolder = ""

if (Test-Path $global:TrackingFilePath) {
    $global:LastProcessedFolder = (Get-Content $global:TrackingFilePath -Raw).Trim()
    Write-Host "Previous progress detected." -ForegroundColor Yellow
    Write-Host "Last processed folder: $global:LastProcessedFolder" -ForegroundColor Yellow
    $resumeChoice = Read-Host "Resume from last position? (Yes/No)"
    if ($resumeChoice -match "^[Yy]") {
        $global:ResumeMode = $true
        Write-Host "Resuming - will skip already processed folders..." -ForegroundColor Green
    } else {
        $global:ResumeMode = $false
        $global:LastProcessedFolder = ""
        Write-Host "Starting fresh run..." -ForegroundColor Green
    }
    Write-Host ""
}

# --- Batch Payload Variables ---
$global:DocumentBatch = @()
$global:BatchSizeEstimate = 0
$global:MaxBatchSize = 24 * 1024 * 1024  # 24 MB
$global:LastBatchFolder = ""
$global:BatchNumber = 0

# --- Debug Mode Counters (reset per company) ---
$global:DbgFilesRead         = 0
$global:DbgFullDocuments     = 0
$global:DbgAttachmentsRead   = 0
$global:DbgAttachmentOnly    = 0
# Function to convert image to base64 data URI
function Convert-ImageToBase64 {
    param([string]$ImagePath)
    try {
        if (-not (Test-Path $ImagePath)) { return $null }
        $bytes = [System.IO.File]::ReadAllBytes($ImagePath)
        $base64 = [System.Convert]::ToBase64String($bytes)
        $mimeType = "image/png"
        if ($bytes.Length -ge 8) {
            if ($bytes[0] -eq 0x89 -and $bytes[1] -eq 0x50 -and $bytes[2] -eq 0x4E -and $bytes[3] -eq 0x47) { $mimeType = "image/png" }
            elseif ($bytes[0] -eq 0xFF -and $bytes[1] -eq 0xD8 -and $bytes[2] -eq 0xFF) { $mimeType = "image/jpeg" }
            elseif ($bytes[0] -eq 0x47 -and $bytes[1] -eq 0x49 -and $bytes[2] -eq 0x46) { $mimeType = "image/gif" }
            elseif ($bytes[0] -eq 0x42 -and $bytes[1] -eq 0x4D) { $mimeType = "image/bmp" }
        }
        return "data:$mimeType;base64,$base64"
    } catch {
        Write-Warning "Failed to convert image: $ImagePath - $($_.Exception.Message)"
        return $null
    }
}

# Process HTML and embed local images as Base64
function Convert-HtmlAndImages {
    param([string]$HtmlFilePath)
    $htmlContent = Get-Content -Path $HtmlFilePath -Raw -Encoding UTF8
    $baseDir = [string](Split-Path $HtmlFilePath -Parent)
    
    # Simple regex to catch image sources.
    $imagePattern = 'src="([^"]+)"'
    $matchesText = [regex]::Matches($htmlContent, $imagePattern)
    
    foreach ($match in $matchesText) {
        $originalPath = [string]$match.Groups[1].Value
        
        # Skip absolute data uris and web links
        if ($originalPath -match '^data:' -or $originalPath -match '^https?://') { continue }
        
        $imagePath = Join-Path $baseDir $originalPath
        $dataUri = Convert-ImageToBase64 -ImagePath $imagePath
        if ($dataUri) {
            $search = "src=`"$originalPath`""
            $replace = "src=`"$dataUri`""
            $htmlContent = $htmlContent.Replace($search, $replace)
        }
    }
    return $htmlContent
}

# Estimate the size of a single document hashtable (in bytes) without serialization
function Get-DocumentSizeEstimate {
    param([hashtable]$Document)
    $size = 200  # JSON structural overhead (keys, braces, commas)
    $size += $Document.companyname.Length
    $size += $Document.documentname.Length
    $size += $Document.documentcontent.Length
    foreach ($att in $Document.attachments) {
        $size += $att.filename.Length + $att.contenttype.Length + $att.content.Length + 100
    }
    return $size
}

# Log an oversized document to the per-company markdown report and the oversized tracking file
function Write-OversizedDocument {
    param(
        [string]$CompanyName,
        [string]$DocumentName,
        [string]$FolderPath,
        [long]$DocumentSizeBytes
    )

    $sizeMB    = [math]::Round($DocumentSizeBytes / 1MB, 2)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

    # Per-company human-readable markdown report
    $mdPath = Join-Path $SourcePath "$CompanyName.md"
    if (-not (Test-Path $mdPath)) {
        $header = @"
# Skipped Documents — $CompanyName

The following documents exceeded the $([math]::Round($global:MaxBatchSize / 1MB)) MB batch limit and were not sent to ServiceNow.
They must be processed manually or once the size limit has been increased.

| Document Name | Size | Folder Path | Timestamp |
|---|---|---|---|
"@
        Set-Content -Path $mdPath -Value $header -Encoding UTF8
    }
    Add-Content -Path $mdPath -Value "| $DocumentName | ${sizeMB} MB | $FolderPath | $timestamp |" -Encoding UTF8

    # Machine-readable oversized tracking file — one folder path per line for future resume
    Add-Content -Path $global:OversizedTrackingPath -Value $FolderPath -Encoding UTF8
}

# Write the per-company debug scan report to a markdown file
function Write-DebugReport {
    param(
        [string]$CompanyName,
        [string]$CompanyPath
    )

    $timestamp  = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $reportPath = Join-Path $SourcePath "$CompanyName-debug-report.md"

    $content = @"
# Debug Scan Report — $CompanyName

Generated: $timestamp

## Summary

| Metric | Count |
|---|---|
| Total Files Read | $($global:DbgFilesRead) |
| Full Documents (Knowledge Articles) | $($global:DbgFullDocuments) |
| Total Attachments Read | $($global:DbgAttachmentsRead) |
| Attachment-Only Entries | $($global:DbgAttachmentOnly) |
"@

    Set-Content -Path $reportPath -Value $content -Encoding UTF8
    Write-Host "  [DEBUG] Report written: $reportPath" -ForegroundColor Yellow
}

# Add a document to the current batch; auto-flush if 25MB cap would be exceeded
function Add-DocumentToBatch {
    param(
        [hashtable]$Document,
        [string]$FolderPath
    )
    $docSize = Get-DocumentSizeEstimate -Document $Document

    # Document exceeds the hard 24 MB cap on its own — flush any pending batch first, then
    # send this document in its own isolated request with the exceededsize flag set
    if ($docSize -gt $global:MaxBatchSize) {
        $sizeMB = [math]::Round($docSize / 1MB, 2)
        Write-Host "  -> [LARGE DOC] '$($Document.documentname)' (~${sizeMB} MB) exceeds the $([math]::Round($global:MaxBatchSize / 1MB)) MB limit. Sending as isolated large document batch." -ForegroundColor Magenta
        Write-OversizedDocument -CompanyName $Document.companyname -DocumentName $Document.documentname -FolderPath $FolderPath -DocumentSizeBytes $docSize
        if ($global:DocumentBatch.Count -gt 0) { Send-CurrentBatch }
        Send-LargeDocument -Document $Document -FolderPath $FolderPath
        return
    }

    # If adding this document would exceed the cap, flush current batch first
    if ($global:DocumentBatch.Count -gt 0 -and ($global:BatchSizeEstimate + $docSize) -gt $global:MaxBatchSize) {
        Send-CurrentBatch
    }

    $global:DocumentBatch += @($Document)
    $global:BatchSizeEstimate += $docSize
    $global:LastBatchFolder = $FolderPath
    Write-Host "  -> Queued into batch (est. ~$([math]::Round($global:BatchSizeEstimate / 1KB)) KB / $([math]::Round($global:MaxBatchSize / 1MB)) MB)" -ForegroundColor DarkGray
}

# Send the current batch of documents to ServiceNow
function Send-CurrentBatch {
    if ($global:DocumentBatch.Count -eq 0) { return }

    $global:BatchNumber++
    $payload = @{ document = $global:DocumentBatch }
    $json = ConvertTo-Json -InputObject $payload -Depth 10 -Compress

    Write-Host ""
    Write-Host "=== Sending Batch #$($global:BatchNumber) ==="  -ForegroundColor Cyan
    Write-Host "Documents : $($global:DocumentBatch.Count)"
    Write-Host "JSON Size : $([math]::Round($json.Length / 1KB)) KB"
    Write-Host "Endpoint  : $global:Endpoint"

    try {
        # Build auth header
        $base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f $global:ServiceNowUser, $global:ServiceNowPass)))

        # Set proper headers
        $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
        $headers.Add('Authorization',('Basic {0}' -f $base64AuthInfo))
        $headers.Add('Accept','application/json')

        # Send HTTP request (UTF-8 encoded body)
        $bodyBytesUtf8 = [System.Text.Encoding]::UTF8.GetBytes($json)
        $bodyBytes = [System.Text.Encoding]::UTF8.GetString($bodyBytesUtf8)
        $response = Invoke-RestMethod -Headers $headers -Method Post -Uri $global:Endpoint -Body $bodyBytes -ContentType "application/json; charset=utf-8"

        Write-Host "[BATCH SEND SUCCESS]" -ForegroundColor Green
        $response | Out-String | Write-Host

        # Update tracking file with the last folder in this batch
        if ($global:LastBatchFolder) {
            $global:LastBatchFolder | Set-Content -Path $global:TrackingFilePath -NoNewline
        }
    } catch {
        Write-Error "Failed to POST batch to ServiceNow Endpoint: $($_.Exception.Message)"
    }

    # Reset batch
    $global:DocumentBatch = @()
    $global:BatchSizeEstimate = 0
    $global:LastBatchFolder = ""
}

# Send a single oversized document as its own isolated batch with X-Exceeded-Payload header
function Send-LargeDocument {
    param(
        [hashtable]$Document,
        [string]$FolderPath
    )

    $global:BatchNumber++
    $payload = @{ document = @($Document) }
    $json = ConvertTo-Json -InputObject $payload -Depth 10 -Compress

    Write-Host ""
    Write-Host "=== Sending Large Document Batch #$($global:BatchNumber) ===" -ForegroundColor Magenta
    Write-Host "Document  : $($Document.documentname)"
    Write-Host "Company   : $($Document.companyname)"
    Write-Host "JSON Size : $([math]::Round($json.Length / 1KB)) KB"
    Write-Host "Endpoint  : $global:Endpoint"

    try {
        $base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f $global:ServiceNowUser, $global:ServiceNowPass)))
        $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
        $headers.Add('Authorization', ('Basic {0}' -f $base64AuthInfo))
        $headers.Add('Accept', 'application/json')
        $headers.Add('X-Exceeded-Payload', 'true')

        $bodyBytesUtf8 = [System.Text.Encoding]::UTF8.GetBytes($json)
        $bodyBytes = [System.Text.Encoding]::UTF8.GetString($bodyBytesUtf8)
        $response = Invoke-RestMethod -Headers $headers -Method Post -Uri $global:Endpoint -Body $bodyBytes -ContentType "application/json; charset=utf-8"

        Write-Host "[LARGE DOC SEND SUCCESS]" -ForegroundColor Green
        $response | Out-String | Write-Host

        # Update tracking file so a resume skips this folder
        if ($FolderPath) {
            $FolderPath | Set-Content -Path $global:TrackingFilePath -NoNewline
        }
    } catch {
        Write-Error "Failed to POST large document to ServiceNow Endpoint: $($_.Exception.Message)"
    }
}

# Check if a company exists in ServiceNow via the check_company endpoint
function Test-CompaniesExist {
    param([string[]]$CompanyNames)
    $checkUrl = "$global:BaseUrl/api/wemop/itglue_knowledge_import/check_company"
    $body = ConvertTo-Json -InputObject @{ companyNames = [array]$CompanyNames } -Compress

    try {
        $base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f $global:ServiceNowUser, $global:ServiceNowPass)))
        $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
        $headers.Add('Authorization',('Basic {0}' -f $base64AuthInfo))
        $headers.Add('Accept','application/json')

        $response = Invoke-RestMethod -Headers $headers -Method Post -Uri $checkUrl -Body $body -ContentType "application/json"

        $resultArray = $response.result.companyNames
        $map = @{}
        for ($i = 0; $i -lt $CompanyNames.Count; $i++) {
            $map[$CompanyNames[$i]] = [bool]$resultArray[$i]
        }
        return $map
    } catch {
        Write-Error "Failed to check companies: $($_.Exception.Message)"
        $map = @{}
        foreach ($name in $CompanyNames) { $map[$name] = $false }
        return $map
    }
}

# Get files from attachments folder, returning Base64-encoded content for JSON binary preservation
function Get-DocumentAttachments {
    param([string]$DocumentId)
    $attFolder = Join-Path $global:AttachmentsRoot "documents\$DocumentId"
    $list = @()
    if (Test-Path $attFolder) {
        foreach ($file in (Get-ChildItem -Path $attFolder -File)) {
            $extension = $file.Extension.TrimStart('.')

            $mimeType = switch ($extension.ToLower()) {
                "pdf"  { "application/pdf" }
                "png"  { "image/png" }
                "jpg"  { "image/jpeg" }
                "jpeg" { "image/jpeg" }
                "gif"  { "image/gif" }
                "bmp"  { "image/bmp" }
                "doc"  { "application/msword" }
                "docx" { "application/vnd.openxmlformats-officedocument.wordprocessingml.document" }
                "xls"  { "application/vnd.ms-excel" }
                "xlsx" { "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }
                "ppt"  { "application/vnd.ms-powerpoint" }
                "pptx" { "application/vnd.openxmlformats-officedocument.presentationml.presentation" }
                "txt"  { "text/plain" }
                "csv"  { "text/csv" }
                "html" { "text/html" }
                "htm"  { "text/html" }
                "xml"  { "application/xml" }
                "json" { "application/json" }
                "zip"  { "application/zip" }
                default { "application/octet-stream" }
            }
            
            try {
                $rawBytes = [System.IO.File]::ReadAllBytes($file.FullName)
                $base64String = [System.Convert]::ToBase64String($rawBytes)

                $list += @{
                    filename    = $file.Name
                    contenttype = $mimeType
                    content     = $base64String
                }
            } catch {
                Write-Warning "Could not read attachment file: $($file.FullName)"
            }
        }
    }
    return $list
}

# Recursive logic for Logic Condition A, B, and C
function Invoke-FolderScan {
    param (
        [string]$FolderPath,
        [string]$CompanyName
    )

    $items = Get-ChildItem -Path $FolderPath -Directory
    foreach ($folder in $items) {
        
        # Skip the root attachments folder entirely during scanning
        if ($folder.FullName -eq $global:AttachmentsRoot) { continue }

        $folderName = $folder.Name

        if ($folderName -match "^DOC-") {
            # Resume mode: skip already-processed DOC- folders
            if ($global:ResumeMode) {
                if ($folder.FullName -eq $global:LastProcessedFolder) {
                    # Found the last processed folder - turn off resume mode, skip this one
                    Write-Host "[SKIP - Already processed] $folderName" -ForegroundColor DarkGray
                    $global:ResumeMode = $false
                    continue
                } else {
                    Write-Host "[SKIP - Already processed] $folderName" -ForegroundColor DarkGray
                    continue
                }
            }

            # Logic A or C
            # Parse DOC-{customer_id}-{document_id} {document_name}
            $matchRegexExt = "^DOC-(\d+)-(\d+)\s+(.+?)(?:\.(\w+))?$"
            $match = [regex]::Match($folderName, $matchRegexExt)

            if ($match.Success) {
                $customerId = $match.Groups[1].Value
                $documentId = $match.Groups[2].Value
                $documentName = $match.Groups[3].Value
                
                $htmlFiles = @(Get-ChildItem -Path $folder.FullName -Filter "*.html")
                $isAttachmentOnly = $false

                # Condition C checks: Ends with '.' and letters (file format), and contains 0kb HTML
                if ($folderName -match "\.[a-zA-Z0-9]+$" -and $htmlFiles.Count -gt 0) {
                    $zeroKbHtml = $htmlFiles | Where-Object { $_.Length -eq 0 }
                    if ($zeroKbHtml) {
                        $isAttachmentOnly = $true
                    }
                }

                if ($isAttachmentOnly) {
                    # Scenario C: Attachment only
                    Write-Host "Found Attachment-Only: $folderName" -ForegroundColor Yellow
                    $attachments = @(Get-DocumentAttachments -DocumentId $documentId)

                    if ($global:DebugMode) {
                        $global:DbgAttachmentOnly++
                        $global:DbgAttachmentsRead += $attachments.Count
                        $global:DbgFilesRead       += $attachments.Count
                    }

                    $doc = @{
                        companyname     = $CompanyName
                        documentname    = $documentName
                        documentcontent = ""
                        attachments     = $attachments
                    }
                    Add-DocumentToBatch -Document $doc -FolderPath $folder.FullName
                }
                else {
                    # Scenario A: Knowledge Article
                    # Condition: Open folder -> contains a non-zero HTML file.
                    # Note: The {customer_id} subfolder is NOT required — it only exists when the
                    # original article had inline images. Articles without it are still valid and
                    # must be processed; Convert-HtmlAndImages handles missing image paths gracefully.
                    $mainHtml = $htmlFiles | Where-Object { $_.Length -gt 0 } | Sort-Object Length -Descending | Select-Object -First 1
                    if ($null -ne $mainHtml) {
                        Write-Host "Found Knowledge Article: $folderName" -ForegroundColor Green

                        try {
                            Write-Host "  -> Parsing HTML: $($mainHtml.Name) ..." -ForegroundColor DarkGray
                            $htmlContent = Convert-HtmlAndImages -HtmlFilePath $mainHtml.FullName

                            Write-Host "  -> Gathering Attachments ..." -ForegroundColor DarkGray
                            $attachments = @(Get-DocumentAttachments -DocumentId $documentId)

                            if ($global:DebugMode) {
                                $global:DbgFullDocuments++
                                $global:DbgAttachmentsRead += $attachments.Count
                                $global:DbgFilesRead       += 1 + $attachments.Count  # 1 HTML + N attachments
                            }

                            $doc = @{
                                companyname     = $CompanyName
                                documentname    = $documentName
                                documentcontent = $htmlContent
                                attachments     = $attachments
                            }
                            Add-DocumentToBatch -Document $doc -FolderPath $folder.FullName
                        } catch {
                            Write-Host "  -> [ERROR] Failed processing article $folderName !" -ForegroundColor Red
                            Write-Error $_.Exception.Message
                        }
                    }
                }
            } else {
                 Write-Warning "DOC- folder didn't match ID regex: $folderName"
            }
        }
        else {
            # Scenario B: Non-DOC subfolder - recurse with current CompanyName (already validated)
            Invoke-FolderScan -FolderPath $folder.FullName -CompanyName $CompanyName
        }
    }
}

try {
    $documentsDir = Join-Path $SourcePath "documents"
    $scanRoot = if (Test-Path $documentsDir) { $documentsDir } else { $SourcePath }

    # --- Phase 1: Validate which companies exist in ServiceNow ---
    Write-Host ""
    Write-Host "=== Phase 1: Validating Companies ===" -ForegroundColor Cyan

    $validCompanies = [System.Collections.Generic.List[hashtable]]::new()
    $companyFolders = Get-ChildItem -Path $scanRoot -Directory |
        Where-Object { $_.Name -notmatch "^DOC-" -and $_.FullName -ne $global:AttachmentsRoot }

    $companyNames = @($companyFolders | ForEach-Object { $_.Name })
    Write-Host "Sending bulk company check for $($companyNames.Count) company(ies)..." -ForegroundColor Cyan

    $companyCheckResults = Test-CompaniesExist -CompanyNames $companyNames

    foreach ($companyFolder in $companyFolders) {
        if ($companyCheckResults[$companyFolder.Name]) {
            Write-Host "  $($companyFolder.Name) [EXISTS]" -ForegroundColor Green
            $validCompanies.Add(@{ Name = $companyFolder.Name; Path = $companyFolder.FullName })
        } else {
            Write-Host "  $($companyFolder.Name) [NOT FOUND - SKIPPED]" -ForegroundColor Red
        }
    }

    Write-Host ""
    Write-Host "$($validCompanies.Count) company(ies) confirmed. Starting document processing..." -ForegroundColor Cyan

    # --- Phase 2: Process documents for each valid company ---
    Write-Host ""
    Write-Host "=== Phase 2: Processing Documents ===" -ForegroundColor Cyan

    foreach ($company in $validCompanies) {
        $global:AttachmentsRoot = Join-Path $company.Path "attachments"
        Write-Host ""
        Write-Host "--- Company: $($company.Name) ---" -ForegroundColor Cyan
        Write-Host "    Attachments root: $($global:AttachmentsRoot)" -ForegroundColor DarkGray

        if ($global:DebugMode) {
            $global:DbgFilesRead       = 0
            $global:DbgFullDocuments   = 0
            $global:DbgAttachmentsRead = 0
            $global:DbgAttachmentOnly  = 0
        }

        Invoke-FolderScan -FolderPath $company.Path -CompanyName $company.Name

        if ($global:DebugMode) {
            Write-DebugReport -CompanyName $company.Name -CompanyPath $company.Path
        }
    }

    # Flush any remaining documents in the batch
    Send-CurrentBatch

    # Clean up tracking file on successful completion
    if (Test-Path $global:TrackingFilePath) {
        Remove-Item $global:TrackingFilePath -Force
        Write-Host "Tracking file removed (full batch completed)." -ForegroundColor DarkGray
    }
    Write-Host "Batch Processing Complete." -ForegroundColor Green
} catch {
    Write-Host ""
    Write-Host "[FATAL SCRIPT ERROR]" -ForegroundColor Red
    Write-Error $_.Exception.Message
    Write-Host "Position: $($_.InvocationInfo.PositionMessage)" -ForegroundColor Red
} finally {
    Write-Host ""
    Read-Host "Press Enter to exit"
}

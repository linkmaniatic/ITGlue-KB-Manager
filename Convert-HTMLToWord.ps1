<#
.SYNOPSIS
    Batch converts HTML documents with embedded images to Word documents.

.DESCRIPTION
    This script searches for folders starting with "DOC-", finds HTML files within them,
    and converts them to Microsoft Word documents. Images referenced in the HTML 
    (which may not have extensions) are converted to base64 and embedded.
    Output files are saved to a specific "ITGlue to Docs" folder.

.PARAMETER SourcePath
    Path to the root directory containing the DOC-* folders. 
    Defaults to the current script directory.

.PARAMETER DestinationPath
    Path where the Word documents will be saved.
    Defaults to "ITGlue to Docs" inside the SourcePath.

.EXAMPLE
    .\Convert-HTMLToWord.ps1
    
.EXAMPLE
    .\Convert-HTMLToWord.ps1 -SourcePath "C:\Downloads\ITGlueExports"
#>

param(
    [Parameter(Mandatory=$false)]
    [string]$SourcePath,
    
    [Parameter(Mandatory=$false)]
    [string]$DestinationPath
)

# Function to get the script directory
function Get-ScriptDirectory {
    if ($PSScriptRoot) {
        return $PSScriptRoot
    }
    else {
        return Split-Path -Parent $MyInvocation.MyCommand.Path
    }
}

# Function to convert image to base64 data URI
function Convert-ImageToBase64 {
    param(
        [string]$ImagePath
    )
    
    try {
        if (-not (Test-Path $ImagePath)) {
            return $null
        }

        $bytes = [System.IO.File]::ReadAllBytes($ImagePath)
        $base64 = [System.Convert]::ToBase64String($bytes)
        
        # Determine MIME type based on file signature
        $mimeType = "image/png"  # Default fallback
        
        if ($bytes.Length -ge 8) {
            # PNG signature
            if ($bytes[0] -eq 0x89 -and $bytes[1] -eq 0x50 -and $bytes[2] -eq 0x4E -and $bytes[3] -eq 0x47) {
                $mimeType = "image/png"
            }
            # JPEG signature
            elseif ($bytes[0] -eq 0xFF -and $bytes[1] -eq 0xD8 -and $bytes[2] -eq 0xFF) {
                $mimeType = "image/jpeg"
            }
            # GIF signature
            elseif ($bytes[0] -eq 0x47 -and $bytes[1] -eq 0x49 -and $bytes[2] -eq 0x46) {
                $mimeType = "image/gif"
            }
            # BMP signature
            elseif ($bytes[0] -eq 0x42 -and $bytes[1] -eq 0x4D) {
                $mimeType = "image/bmp"
            }
        }
        
        return "data:$mimeType;base64,$base64"
    }
    catch {
        Write-Warning "Failed to convert image: $ImagePath - $($_.Exception.Message)"
        return $null
    }
}

# Initialize paths
if (-not $SourcePath) {
    $SourcePath = Get-ScriptDirectory
}
if (-not [System.IO.Path]::IsPathRooted($SourcePath)) {
    $SourcePath = Join-Path (Get-ScriptDirectory) $SourcePath
}

if (-not $DestinationPath) {
    $DestinationPath = Join-Path $SourcePath "ITGlue to Docs"
}
if (-not [System.IO.Path]::IsPathRooted($DestinationPath)) {
    $DestinationPath = Join-Path (Get-ScriptDirectory) $DestinationPath
}

# Create destination directory if it doesn't exist
if (-not (Test-Path $DestinationPath)) {
    New-Item -ItemType Directory -Path $DestinationPath -Force | Out-Null
    Write-Host "Created output directory: $DestinationPath" -ForegroundColor Cyan
}

# Initialize Word application once
try {
    Write-Host "Initializing Word Application..." -ForegroundColor Cyan
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
}
catch {
    Write-Error "Failed to initialize Microsoft Word. Please ensure it is installed."
    exit 1
}

try {
    # Find all DOC-* folders
    Write-Host "Scanning for DOC-* folders in: $SourcePath"
    $docFolders = Get-ChildItem -Path $SourcePath -Directory -Filter "DOC-*" -Recurse
    
    if ($docFolders.Count -eq 0) {
        Write-Warning "No folders starting with 'DOC-' found in source path."
    }

    foreach ($folder in $docFolders) {
        Write-Host "Processing folder: $($folder.Name)" -ForegroundColor Cyan
        
        # Find HTML files in this folder
        $htmlFiles = Get-ChildItem -Path $folder.FullName -Filter "*.html"
        
        foreach ($htmlFile in $htmlFiles) {
            Write-Host "  Found Document: $($htmlFile.Name)"
            
            $outputFileName = [System.IO.Path]::ChangeExtension([string]$htmlFile.Name, ".docx")
            $outputFilePath = Join-Path $DestinationPath $outputFileName
            
            # Check if file already exists
            if (Test-Path $outputFilePath) {
                Write-Host "    File already exists in output folder. Skipping..." -ForegroundColor Yellow
                continue
            }
            
            try {
                # --- Conversion Logic ---
                $baseDir = [string]$htmlFile.DirectoryName
                $fullHtmlPath = [string]$htmlFile.FullName
                $htmlContent = Get-Content -Path $fullHtmlPath -Raw -Encoding UTF8
                
                # Image processing
                $imagePattern = 'src="([^"]+)"'
                $matchesText = [regex]::Matches($htmlContent, $imagePattern)
                
                if ($matchesText.Count -gt 0) {
                     Write-Host "    Processing $($matchesText.Count) images..." -ForegroundColor Gray
                }
                
                foreach ($match in $matchesText) {
                    $originalPath = [string]$match.Groups[1].Value
                    
                    if ($originalPath -match '^data:' -or $originalPath -match '^https?://') { continue }
                    
                    # Construct full image path
                    $imagePath = Join-Path $baseDir $originalPath
                    
                    $dataUri = Convert-ImageToBase64 -ImagePath $imagePath
                    if ($dataUri) {
                        # Explicitly cast to string for .NET Replace method
                        $search = "src=`"$originalPath`""
                        $replace = "src=`"$dataUri`""
                        $htmlContent = $htmlContent.Replace($search, $replace)
                    }
                }
                
                # Add HTML structure if missing
                if ($htmlContent -notmatch '(?i)<html[\s>]') {
                    $htmlContent = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Document</title>
    <style>
        body { font-family: Arial, sans-serif; font-size: 11pt; line-height: 1.5; margin: 1in; }
        h1 { font-size: 18pt; font-weight: bold; margin: 12pt 0 6pt 0; }
        h2 { font-size: 14pt; font-weight: bold; margin: 10pt 0 6pt 0; }
        img { max-width: 100%; height: auto; }
        a { color: #0563C1; text-decoration: underline; }
    </style>
</head>
<body>
$htmlContent
</body>
</html>
"@
                }
                
                # Save temp HTML
                $tempHtmlPath = [System.IO.Path]::GetTempFileName()
                $tempHtmlPath = [System.IO.Path]::ChangeExtension($tempHtmlPath, ".html")
                $htmlContent | Out-File -FilePath $tempHtmlPath -Encoding UTF8
                
                # Open in Word
                $doc = $word.Documents.Open([string]$tempHtmlPath, $false, $false)
                
                # Save as DOCX
                $outputFilePath = [string]$outputFilePath
                $doc.SaveAs([ref]$outputFilePath, [ref]16) # wdFormatDocumentDefault
                $doc.Close([ref]$false)
                
                # Cleanup temp file
                Remove-Item $tempHtmlPath -Force -ErrorAction SilentlyContinue
                
                Write-Host "    Converted -> $outputFileName" -ForegroundColor Green
            }
            catch {
                Write-Error "    Failed to convert $($htmlFile.Name): $($_.Exception.Message)"
                Write-Error "    Stack Trace: $($_.Exception.StackTrace)"
            }
        }
    }
}
finally {
    if ($word) {
        $word.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
    }
    Remove-Variable word -ErrorAction SilentlyContinue
    
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

Write-Host ""
Write-Host "Batch Processing Complete." -ForegroundColor Green
if (Test-Path $DestinationPath) {
    Invoke-Item $DestinationPath
}

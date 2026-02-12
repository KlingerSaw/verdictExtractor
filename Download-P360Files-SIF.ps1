# P360 SIF API Document Downloader
# Henter dokumenter via P360 SIF REST API

# Set console encoding to UTF-8 for Danish characters
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

param(
    [ValidateSet("arkiv", "produktion")][string]$Environment,
    [ValidateSet("word", "pdf", "both")][string]$FileType,
    [string]$OutputDir,
    [string]$MarkdownDir,
    [string]$AuthKey,
    [int]$ContactRecno,
    [string]$TitleFilter
)

# Set defaults
if (-not $Environment) { $Environment = "produktion" }
if (-not $FileType) { $FileType = "both" }
if (-not $OutputDir) { $OutputDir = ".\prod_downloads" }
if (-not $MarkdownDir) { $MarkdownDir = ".\prod_markdown" }
if (-not $ContactRecno) { $ContactRecno = 100016 }
if (-not $TitleFilter) { $TitleFilter = "Afgørelse af%" }

Write-Host ""
Write-Host "====================================================================" -ForegroundColor Cyan
Write-Host " P360 SIF API Document Downloader" -ForegroundColor Cyan
Write-Host "====================================================================" -ForegroundColor Cyan
Write-Host ""

# Mode selection
Write-Host "Vaelg funktion:" -ForegroundColor Yellow
Write-Host "  [1] Download filer og konverter til markdown" -ForegroundColor Cyan
Write-Host "  [2] Konverter eksisterende filer til markdown (skip download)" -ForegroundColor Cyan
$modeChoice = Read-Host "Vaelg (1-2)"
$convertOnly = ($modeChoice -eq '2')

if ($convertOnly) {
    Write-Host ""
    Write-Host "[*] Mode: KUN KONVERTERING (springer download over)" -ForegroundColor Yellow
    Write-Host ""
}

if (-not $convertOnly) {
    # Configuration
    $baseUrl = if ($Environment -eq 'arkiv') { 
        'https://esdh-nh-arkiv/Biz/v2/api/call/SI.Data.RPC/SI.Data.RPC' 
    } else { 
        'https://esdh-nh-PB360/Biz/v2/api/call/SI.Data.RPC/SI.Data.RPC' 
    }

    $downloadBaseUrl = if ($Environment -eq 'arkiv') { 
        'https://esdh-nh-arkiv' 
    } else { 
        'https://esdh-nh-pb360' 
    }

    # Prompt for AuthKey if not provided
    if (-not $AuthKey) {
        $AuthKey = Read-Host "P360 AuthKey"
    }

    # Prompt for filetype if not specified
    if (-not $FileType -or $FileType -eq 'both') {
        Write-Host ""
        Write-Host "Vaelg filtype:" -ForegroundColor Yellow
        Write-Host "  [1] Word (DOC/DOCX)" -ForegroundColor Cyan
        Write-Host "  [2] PDF" -ForegroundColor Cyan
        Write-Host "  [3] Begge (Word + PDF)" -ForegroundColor Cyan
        $typeChoice = Read-Host "Vaelg (1-3)"
        if ($typeChoice -eq '1') { $FileType = 'word' }
        elseif ($typeChoice -eq '2') { $FileType = 'pdf' }
        else { $FileType = 'both' }
    }

    # Mask AuthKey for display
    $maskedKey = if ($AuthKey.Length -gt 8) { 
        $AuthKey.Substring(0, 8) + "****" 
    } else { 
        "****" 
    }

    Write-Host ""
    Write-Host "[+] Miljoe: $Environment" -ForegroundColor Green
    Write-Host "[+] Base URL: $baseUrl" -ForegroundColor Green
    Write-Host "[+] AuthKey: $maskedKey" -ForegroundColor Green
    Write-Host "[+] ContactRecno: $ContactRecno" -ForegroundColor Green
    Write-Host "[+] Title filter: $TitleFilter" -ForegroundColor Green
    Write-Host "[+] Filtype: $FileType" -ForegroundColor Green
    Write-Host ""
} # End if not convertOnly (config section)

# Create output directories
if (-not (Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir | Out-Null
    Write-Host "[+] Oprettet output mappe: $OutputDir" -ForegroundColor Green
}

if (-not (Test-Path $MarkdownDir)) {
    New-Item -ItemType Directory -Path $MarkdownDir | Out-Null
    Write-Host "[+] Oprettet markdown mappe: $MarkdownDir" -ForegroundColor Green
}

if (-not $convertOnly) {
    Write-Host ""
    Write-Host "====================================================================" -ForegroundColor Cyan
    Write-Host " HENTER DOKUMENTER FRA P360 SIF API" -ForegroundColor Cyan
    Write-Host "====================================================================" -ForegroundColor Cyan
    Write-Host ""

    # Build API request with pagination
$apiUrl = "$baseUrl/DocumentService/GetDocuments?authkey=$AuthKey"

$allDocuments = @()
$page = 0
$hasMorePages = $true

while ($hasMorePages) {
    $requestBody = @{
        parameter = @{
            Page = $page
            IncludeCustomFields = "false"
            ContactRecnos = @($ContactRecno)
            Title = $TitleFilter
        }
    } | ConvertTo-Json -Depth 10

    if ($page -eq 0) {
        Write-Host "[*] Kalder API (side $page)..." -ForegroundColor Yellow
    } else {
        Write-Host "[*] Henter side $page..." -ForegroundColor Yellow
    }

    # Call API
    try {
        $response = Invoke-RestMethod -Uri $apiUrl -Method Post -Body $requestBody -ContentType "application/json"
        
        if ($response.Successful -eq $false) {
            Write-Host "FEJL: API kald fejlede" -ForegroundColor Red
            Write-Host "ErrorMessage: $($response.ErrorMessage)" -ForegroundColor Red
            Write-Host "ErrorDetails: $($response.ErrorDetails)" -ForegroundColor Red
            pause
            exit 1
        }
        
        $pageDocuments = $response.Documents
        
        if ($pageDocuments -and $pageDocuments.Count -gt 0) {
            Write-Host "    Modtaget $($pageDocuments.Count) dokumenter" -ForegroundColor Gray
            $allDocuments += $pageDocuments
            
            # Check if there are more pages (API returns 100 per page)
            if ($pageDocuments.Count -eq 100) {
                $page++
            } else {
                $hasMorePages = $false
            }
        } else {
            $hasMorePages = $false
        }
        
    } catch {
        Write-Host "FEJL: Kunne ikke kalde API" -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
        pause
        exit 1
    }
}

Write-Host ""
Write-Host "[+] TOTALT modtaget $($allDocuments.Count) dokumenter" -ForegroundColor Green

Write-Host ""
Write-Host "[*] Filtrerer dokumenter (KUN DOC/DOCX/PDF + Afgørelse af)..." -ForegroundColor Yellow

# Filter documents
$allowedExtensions = @('DOC', 'DOCX', 'PDF')
if ($FileType -eq 'word') { 
    $allowedExtensions = @('DOC', 'DOCX') 
} elseif ($FileType -eq 'pdf') { 
    $allowedExtensions = @('PDF') 
}

$filesToDownload = @()
$skipped = 0
$skipReasons = @{}

foreach ($doc in $allDocuments) {
    $docRecno = $doc.Recno
    $docNumber = $doc.DocumentNumber
    $docTitle = $doc.Title
    $caseNumber = $doc.CaseNumber
    $caseRecno = $doc.CaseRecno
    $caseTitle = if ($doc.CaseNameAndDescription) { $doc.CaseNameAndDescription } else { "" }
    $klassifikation = if ($doc.AccessCodeCode) { $doc.AccessCodeCode } else { "" }
    
    # Build P360 links
    $documentLink = "$downloadBaseUrl/Biz?action=OpenDocument&documentRecno=$docRecno"
    $caseLink = "$downloadBaseUrl/Biz?action=OpenCase&caseRecno=$caseRecno"
    
    # Document-level validation BEFORE processing files
    $skipDoc = $false
    $skipReason = ""
    
    # Rule 1: Must start with "Afgørelse af"
    if ($docTitle -notmatch '^Afg.relse af') {
        $skipDoc = $true
        $skipReason = "Titel!=Afgørelse af"
    }
    
    # Rule 2: Check klassifikation
    if (-not $skipDoc -and $klassifikation -match '2100|EFTERLEVELSE') {
        $skipDoc = $true
        $skipReason = "Klassifikation=Efterlevelse"
    }
    
    # Rule 3: Check case title
    if (-not $skipDoc -and $caseTitle -match 'EFTERLEVELSE|OMKOSTNINGSDAEKNING') {
        $skipDoc = $true
        $skipReason = "Sagstitel=Ekskluderet"
    }
    
    # Rule 4: Check document title for EFTERLEVELSE
    if (-not $skipDoc -and $docTitle -match 'EFTERLEVELSE') {
        $skipDoc = $true
        $skipReason = "Dokumenttitel=Efterlevelse"
    }
    
    if ($skipDoc) {
        $skipped++
        if (-not $skipReasons.ContainsKey($skipReason)) {
            $skipReasons[$skipReason] = 0
        }
        $skipReasons[$skipReason]++
        
        if ($skipped -le 10) {
            Write-Host "o SKIP: $skipReason | Titel='$docTitle'" -ForegroundColor DarkGray
        }
        continue
    }
    
    # Process each file in document
    if ($doc.Files -and $doc.Files.Count -gt 0) {
        foreach ($file in $doc.Files) {
            $fileRecno = $file.Recno
            $fileTitle = $file.Title
            $fileFormat = $file.Format.ToUpper()
            $fileUrl = $file.URL
            
            # Validate extension
            if ($fileFormat -notin $allowedExtensions) {
                $skipped++
                $reason = "Extension=$fileFormat"
                if (-not $skipReasons.ContainsKey($reason)) {
                    $skipReasons[$reason] = 0
                }
                $skipReasons[$reason]++
                
                if ($skipped -le 10) {
                    Write-Host "o SKIP: $reason | Fil='$fileTitle'" -ForegroundColor DarkGray
                }
                continue
            }
            
            Write-Host "v OK: '$fileTitle' | Format=$fileFormat | FileRecno=$fileRecno" -ForegroundColor Green
            
            $filesToDownload += @{
                FileRecno = $fileRecno
                Filename = $fileTitle
                Format = $fileFormat
                URL = $fileUrl
                DocumentRecno = $docRecno
                DocumentNumber = $docNumber
                DocumentTitle = $docTitle
                CaseNumber = $caseNumber
                CaseRecno = $caseRecno
                DocumentLink = $documentLink
                CaseLink = $caseLink
            }
        }
    } else {
        $skipped++
        $reason = "Ingen filer"
        if (-not $skipReasons.ContainsKey($reason)) {
            $skipReasons[$reason] = 0
        }
        $skipReasons[$reason]++
        
        if ($skipped -le 10) {
            Write-Host "o SKIP: $reason | Titel='$docTitle'" -ForegroundColor DarkGray
        }
    }
}

Write-Host ""
Write-Host "[+] $($filesToDownload.Count) filer klar til download" -ForegroundColor Green
Write-Host "[+] $skipped dokumenter/filer sprunget over:" -ForegroundColor Yellow
foreach ($reason in $skipReasons.Keys | Sort-Object) {
    Write-Host "    - $reason`: $($skipReasons[$reason])" -ForegroundColor Gray
}
Write-Host ""

if ($filesToDownload.Count -eq 0) {
    Write-Host "INGEN filer at hente!" -ForegroundColor Red
    pause
    exit 0
}

$confirm = Read-Host "Start download? (Y/N)"
if ($confirm -ne 'Y' -and $confirm -ne 'y') {
    Write-Host "Afbrudt" -ForegroundColor Yellow
    exit 0
}

Write-Host ""
Write-Host "====================================================================" -ForegroundColor Cyan
Write-Host " DOWNLOADER FILER" -ForegroundColor Cyan
Write-Host "====================================================================" -ForegroundColor Cyan
Write-Host ""

$downloaded = 0
$errors = 0
$downloadedFiles = @()

# Use Windows Authentication (no credentials needed if running as domain user)
$session = New-Object Microsoft.PowerShell.Commands.WebRequestSession

foreach ($file in $filesToDownload) {
    $fileNum = $downloaded + $errors + 1
    $total = $filesToDownload.Count
    
    # Sanitize filename
    $safeFilename = $file.Filename -replace '[\\/:*?"<>|]', '_'
    $outputPath = Join-Path $OutputDir $safeFilename
    
    Write-Host "v [$fileNum/$total] '$($file.Filename)' -> Henter..." -ForegroundColor Cyan
    
    try {
        $startTime = Get-Date
        
        # Download using URL from API with Windows Authentication
        Invoke-WebRequest -Uri $file.URL -OutFile $outputPath -UseDefaultCredentials -WebSession $session -UseBasicParsing
        
        $duration = (Get-Date) - $startTime
        $fileSize = (Get-Item $outputPath).Length / 1MB
        Write-Host "v [$fileNum/$total] '$($file.Filename)' -> OK ($("{0:N2}" -f $fileSize) MB, $("{0:N1}" -f $duration.TotalSeconds)s)" -ForegroundColor Green
        $downloaded++
        
        # Add to list for markdown conversion
        $downloadedFiles += @{
            Path = $outputPath
            Filename = $file.Filename
            Format = $file.Format
            DocumentTitle = $file.DocumentTitle
            DocumentNumber = $file.DocumentNumber
            CaseNumber = $file.CaseNumber
            DocumentLink = $file.DocumentLink
            CaseLink = $file.CaseLink
            FileRecno = $file.FileRecno
            DocumentRecno = $file.DocumentRecno
        }
        
    } catch {
        Write-Host "x [$fileNum/$total] '$($file.Filename)' -> FEJL ($($_.Exception.Message))" -ForegroundColor Red
        $errors++
    }
    
    Start-Sleep -Milliseconds 300
}

Write-Host ""
Write-Host "====================================================================" -ForegroundColor Cyan
Write-Host " DOWNLOAD FAERDIG" -ForegroundColor Cyan
Write-Host "====================================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "v $downloaded filer hentet" -ForegroundColor Green
Write-Host "x $errors fejl" -ForegroundColor $(if ($errors -gt 0) { "Red" } else { "Green" })
Write-Host ""

} else {
    # CONVERT-ONLY MODE - Scan existing files in download folder
    Write-Host ""
    Write-Host "====================================================================" -ForegroundColor Cyan
    Write-Host " SCANNER EKSISTERENDE FILER" -ForegroundColor Cyan
    Write-Host "====================================================================" -ForegroundColor Cyan
    Write-Host ""
    
    $downloadedFiles = @()
    
    if (Test-Path $OutputDir) {
        $existingFiles = Get-ChildItem -Path $OutputDir -File | Where-Object { $_.Extension -match '\.(pdf|docx?|doc)$' }
        
        Write-Host "[+] Fundet $($existingFiles.Count) filer i $OutputDir" -ForegroundColor Green
        
        foreach ($file in $existingFiles) {
            # Try to parse filename to extract metadata
            $basename = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
            $extension = $file.Extension.TrimStart('.').ToUpper()
            
            # Try to extract case number from filename (e.g., "20_01453 Afgørelse")
            $caseNumber = ""
            $docNumber = ""
            if ($basename -match '^(\d{2}[_/]\d{5})') {
                $caseNumber = $Matches[1] -replace '_', '/'
                $docNumber = $caseNumber + "-X"  # We don't know the document suffix
            }
            
            $downloadedFiles += @{
                Path = $file.FullName
                DocumentTitle = $basename
                Format = $extension
                DocumentNumber = $docNumber
                CaseNumber = $caseNumber
                Filename = $file.Name
                DocumentLink = ""
                CaseLink = ""
            }
        }
        
        Write-Host "[+] $($downloadedFiles.Count) filer klar til konvertering" -ForegroundColor Green
    } else {
        Write-Host "[!] Download mappe findes ikke: $OutputDir" -ForegroundColor Red
        Write-Host "[!] Opret mappen og placer filer der, eller koer download mode" -ForegroundColor Yellow
        pause
        exit 1
    }
    
    Write-Host ""
}

# Convert files to markdown
if ($downloadedFiles.Count -gt 0) {
    Write-Host ""
    Write-Host "====================================================================" -ForegroundColor Cyan
    Write-Host " KONVERTERER TIL MARKDOWN" -ForegroundColor Cyan
    Write-Host "====================================================================" -ForegroundColor Cyan
    Write-Host ""

    # Find script directory for pdftotext
    $scriptDir = if ($PSScriptRoot) {
        $PSScriptRoot
    } elseif ($MyInvocation.MyCommand.Path) {
        Split-Path -Parent $MyInvocation.MyCommand.Path
    } else {
        Get-Location
    }

    # Check for pdftotext.exe
    $pdfToTextPath = $null
    $searchPaths = @(
        (Join-Path $scriptDir "pdftotext.exe"),
        (Join-Path $scriptDir "bin64\pdftotext.exe"),
        (Join-Path $scriptDir "xpdf-tools\bin64\pdftotext.exe")
    )

    foreach ($path in $searchPaths) {
        if (Test-Path $path) {
            $pdfToTextPath = $path
            Write-Host "[+] Fundet pdftotext: $path" -ForegroundColor Green
            break
        }
    }

    if (-not $pdfToTextPath) {
        Write-Host "[!] pdftotext.exe ikke fundet - PDF'er vil kun have metadata" -ForegroundColor Yellow
    }

    $converted = 0
    $conversionErrors = 0

    foreach ($file in $downloadedFiles) {
        $fileNum = $converted + $conversionErrors + 1
        $total = $downloadedFiles.Count
        
        $inputPath = $file.Path
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($file.Filename)
        $markdownPath = Join-Path $MarkdownDir "$baseName.md"
        
        # Check if already converted
        if (Test-Path $markdownPath) {
            Write-Host "  [$fileNum/$total] $($file.Filename) -> FINDES ALLEREDE (springer over)" -ForegroundColor Yellow
            $converted++  # Count as converted since it exists
            continue
        }
        
        Write-Host "  [$fileNum/$total] $($file.Filename) -> " -NoNewline -ForegroundColor Cyan
        
        try {
            $markdown = ""
            
            if ($file.Format -eq 'PDF') {
                # PDF to Markdown
                $markdown = "# $($file.DocumentTitle)`n`n"
                $markdown += "**Dokument:** $($file.DocumentNumber)`n"
                $markdown += "**Sag:** $($file.CaseNumber)`n"
                $markdown += "**Format:** PDF`n"
                $markdown += "**P360 Links:**`n"
                $markdown += "- [Åbn dokument]($($file.DocumentLink))`n"
                $markdown += "- [Åbn sag]($($file.CaseLink))`n`n"
                $markdown += "---`n`n"
                
                if ($pdfToTextPath) {
                    # Extract text using pdftotext
                    $tempTxt = [System.IO.Path]::GetTempFileName()
                    $absolutePath = (Resolve-Path $inputPath).Path
                    
                    & $pdfToTextPath -layout -enc UTF-8 "$absolutePath" "$tempTxt" 2>$null
                    
                    if (Test-Path $tempTxt) {
                        $text = Get-Content $tempTxt -Raw -Encoding UTF8 -ErrorAction SilentlyContinue
                        if ($text -and $text.Trim().Length -gt 50) {
                            $markdown += $text
                        } else {
                            $markdown += "*[PDF indeholder ingen udtrækbar tekst - muligvis scannet dokument]*"
                        }
                        Remove-Item $tempTxt -Force -ErrorAction SilentlyContinue
                    } else {
                        $markdown += "*[Kunne ikke udtrække tekst fra PDF]*"
                    }
                } else {
                    $markdown += "*[pdftotext ikke tilgængelig - installer for tekstudtræk]*"
                }
                
            } elseif ($file.Format -eq 'DOCX' -or $file.Format -eq 'DOC') {
                # Word to Markdown via Word COM object
                $markdown = "# $($file.DocumentTitle)`n`n"
                $markdown += "**Dokument:** $($file.DocumentNumber)`n"
                $markdown += "**Sag:** $($file.CaseNumber)`n"
                $markdown += "**Format:** $($file.Format)`n"
                if ($file.DocumentLink) {
                    $markdown += "**P360 Links:**`n"
                    $markdown += "- [Åbn dokument]($($file.DocumentLink))`n"
                    $markdown += "- [Åbn sag]($($file.CaseLink))`n`n"
                }
                $markdown += "---`n`n"
                
                try {
                    # Use Word COM object to convert
                    $word = New-Object -ComObject Word.Application
                    $word.Visible = $false
                    
                    $absolutePath = (Resolve-Path $inputPath).Path
                    $doc = $word.Documents.Open($absolutePath)
                    
                    # Extract text from Word document
                    $text = $doc.Content.Text
                    
                    # Close document
                    $doc.Close($false)
                    $word.Quit()
                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null
                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
                    [System.GC]::Collect()
                    [System.GC]::WaitForPendingFinalizers()
                    
                    if ($text -and $text.Trim().Length -gt 50) {
                        $markdown += $text
                    } else {
                        $markdown += "*[Tomt dokument eller kunne ikke udtrække tekst]*"
                    }
                } catch {
                    $markdown += "*[Kunne ikke aabne Word dokument - Microsoft Word skal vaere installeret]*`n"
                    $markdown += "*Fejl: $($_.Exception.Message)*"
                }
            }
            
            # Save markdown with UTF-8 BOM
            $utf8 = New-Object System.Text.UTF8Encoding $true
            [System.IO.File]::WriteAllText($markdownPath, $markdown, $utf8)
            
            Write-Host "OK" -ForegroundColor Green
            $converted++
            
        } catch {
            Write-Host "FEJL" -ForegroundColor Red
            $conversionErrors++
        }
    }

    Write-Host ""
    Write-Host "[+] $converted filer konverteret til markdown" -ForegroundColor Green
    if ($conversionErrors -gt 0) {
        Write-Host "[!] $conversionErrors konverteringer fejlede" -ForegroundColor Yellow
    }
    Write-Host ""
}

# Create markdown index file with links
Write-Host "[*] Opretter markdown index..." -ForegroundColor Yellow

$indexContent = @"
# P360 Dokumenter - $Environment

**Hentet:** $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')  
**Antal filer:** $downloaded  
**Filter:** $TitleFilter

## Dokumenter

"@

foreach ($file in $downloadedFiles) {
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($file.Filename)
    $markdownFile = "$baseName.md"
    
    $indexContent += @"

### $($file.DocumentTitle)

- **Dokument nummer:** $($file.DocumentNumber)
- **Sags nummer:** $($file.CaseNumber)
- **Fil:** ``$($file.Filename)`` ($($file.Format))
- **P360 Links:**
  - [Åbn dokument i P360]($($file.DocumentLink))
  - [Åbn sag i P360]($($file.CaseLink))
- **Lokal fil:** [``$($file.Filename)``](../prod_downloads/$($file.Filename))
- **Markdown:** [``$markdownFile``](./$markdownFile)

"@
}

$indexPath = Join-Path $MarkdownDir "INDEX.md"
$utf8 = New-Object System.Text.UTF8Encoding $true
[System.IO.File]::WriteAllText($indexPath, $indexContent, $utf8)

Write-Host "[+] Markdown index oprettet: $indexPath" -ForegroundColor Green

Write-Host ""
Write-Host "Filer gemt i:" -ForegroundColor Cyan
Write-Host "  Downloads: $OutputDir" -ForegroundColor Cyan
Write-Host "  Index: $indexPath" -ForegroundColor Cyan
Write-Host ""

$openFolder = Read-Host "Aaben markdown mappe? (Y/N)"
if ($openFolder -eq 'Y' -or $openFolder -eq 'y') {
    Invoke-Item $MarkdownDir
}

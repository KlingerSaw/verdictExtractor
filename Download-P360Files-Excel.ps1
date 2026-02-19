# P360 Document Downloader - Excel Edition (ARKIV)
# Læser data fra Excel export i stedet for API

param(
    [string]$ExcelFile,
    [ValidateSet("word", "pdf", "both")][string]$FileType,
    [string]$Username,
    [string]$Password,
    [string]$OutputDir,
    [string]$MarkdownDir,
    [int]$MaxFilesToProcess,
    [int]$RowsToSkip
)

# Set console encoding to UTF-8 for Danish characters
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

function Normalize-MarkdownText {
    param([string]$Text)

    if ($null -eq $Text) {
        return ""
    }

    $normalized = $Text

    # Normalize hidden Unicode chars that often break markdown links and regex matches
    $normalized = $normalized.Replace([char]0x00A0, ' ') # NBSP
    $normalized = $normalized.Replace([string][char]0x200B, '')  # ZWSP
    $normalized = $normalized.Replace([string][char]0x200C, '')  # ZWNJ
    $normalized = $normalized.Replace([string][char]0x200D, '')  # ZWJ
    $normalized = $normalized.Replace([string][char]0xFEFF, '')  # BOM/ZWNBSP

    # Normalize line endings before cleanup
    $normalized = $normalized -replace "`r`n", "`n"
    $normalized = $normalized -replace "`r", "`n"

    # Remove ASCII control characters while preserving tab and newline
    $normalized = $normalized -replace "[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]", ""

    return [string]::Normalize($normalized, [Text.NormalizationForm]::FormC)
}

function Build-MarkdownHeader {
    param(
        [hashtable]$FileInfo,
        [string]$FormatLabel
    )

    $downloadUrl = $null
    $documentUrl = $null
    if ($FileInfo.FileId) {
        $downloadUrl = "https://esdh-nh-arkiv/GetFile.aspx?fileId=$($FileInfo.FileId)&redirect=true"
    }
    if ($FileInfo.DocumentRecno) {
        $documentUrl = "https://esdh-nh-arkiv/locator/Earchive/Case/Details/locator.aspx?name=Earchive.Document.Details.EArchive&module=Document&subtype=17&recno=$($FileInfo.DocumentRecno)"
    }

    $markdown = ""
    $markdown += "# $($FileInfo.Title)`n`n"
    $markdown += "**Dokument:** $($FileInfo.DocumentNumber)`n"
    $markdown += "**Sag:** $($FileInfo.CaseNumber)`n"
    $markdown += "**Format:** $FormatLabel`n"
    $markdown += "**FileID:** $($FileInfo.FileId)`n"

    if ($downloadUrl -or $documentUrl) {
        $markdown += "**P360 Links:**`n"
        if ($downloadUrl) {
            $markdown += "- [Hent fil]($downloadUrl)`n"
        }
        if ($documentUrl) {
            $markdown += "- [Dokumentkort]($documentUrl)`n"
        }
        $markdown += "`n"
    } else {
        $markdown += "`n"
    }

    $markdown += "---`n`n"
    return $markdown
}

function Convert-DownloadedFileToMarkdown {
    param(
        [hashtable]$FileInfo,
        [string]$MarkdownDir,
        [string]$PdfToTextPath
    )

    $inputPath = $FileInfo.Path
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($FileInfo.Filename)
    $markdownPath = Join-Path $MarkdownDir "$baseName.md"

    if (Test-Path $markdownPath) {
        Write-Host "    [MD] Findes allerede: $markdownPath" -ForegroundColor Yellow
        return $true
    }

    $markdown = ""

    if ($FileInfo.Extension -eq 'PDF') {
        $markdown = Build-MarkdownHeader -FileInfo $FileInfo -FormatLabel 'PDF'

        if ($PdfToTextPath) {
            $tempTxt = [System.IO.Path]::GetTempFileName()
            $absolutePath = (Resolve-Path $inputPath).Path
            & $PdfToTextPath -layout -enc UTF-8 "$absolutePath" "$tempTxt" 2>$null

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

    } elseif ($FileInfo.Extension -eq 'DOCX' -or $FileInfo.Extension -eq 'DOC') {
        $markdown = Build-MarkdownHeader -FileInfo $FileInfo -FormatLabel $FileInfo.Extension

        try {
            $word = New-Object -ComObject Word.Application
            $word.Visible = $false

            $absolutePath = (Resolve-Path $inputPath).Path
            $doc = $word.Documents.Open($absolutePath)
            $text = $doc.Content.Text

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

    $utf8 = New-Object System.Text.UTF8Encoding $true
    [System.IO.File]::WriteAllText($markdownPath, (Normalize-MarkdownText -Text $markdown), $utf8)
    Write-Host "    [MD] Oprettet: $markdownPath" -ForegroundColor Green

    return $true
}

function Resolve-PdfToTextPath {
    param([string]$ScriptDir)

    $searchPaths = @(
        (Join-Path $ScriptDir "pdftotext.exe"),
        (Join-Path $ScriptDir "bin64\pdftotext.exe"),
        (Join-Path $ScriptDir "xpdf-tools\bin64\pdftotext.exe")
    )

    foreach ($path in $searchPaths) {
        if (Test-Path $path) {
            Write-Host "[+] Fundet pdftotext: $path" -ForegroundColor Green
            return $path
        }
    }

    Write-Host "[!] pdftotext.exe ikke fundet - PDF'er vil kun have metadata" -ForegroundColor Yellow
    return $null
}

# Set defaults
if (-not $FileType) { $FileType = "both" }
if (-not $OutputDir) { $OutputDir = ".\arkiv_downloads" }
if (-not $MarkdownDir) { $MarkdownDir = ".\arkiv_markdown" }
if (-not $MaxFilesToProcess -or $MaxFilesToProcess -lt 0) { $MaxFilesToProcess = 0 }
if (-not $RowsToSkip -or $RowsToSkip -lt 0) { $RowsToSkip = 0 }

Write-Host ""
Write-Host "====================================================================" -ForegroundColor Cyan
Write-Host " P360 Document Downloader - Excel Edition (ARKIV)" -ForegroundColor Cyan
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


if ($MaxFilesToProcess -le 0) {
    $maxFilesInput = Read-Host "Maks antal filer at behandle? (tryk Enter for alle)"
    if (-not [string]::IsNullOrWhiteSpace($maxFilesInput)) {
        $parsedMaxFiles = 0
        if ([int]::TryParse($maxFilesInput.Trim(), [ref]$parsedMaxFiles) -and $parsedMaxFiles -gt 0) {
            $MaxFilesToProcess = $parsedMaxFiles
        } else {
            Write-Host "[!] Ugyldigt antal - bruger alle filer" -ForegroundColor Yellow
            $MaxFilesToProcess = 0
        }
    }
}

if ($RowsToSkip -gt 0) {
    Write-Host "[+] Raekker der springes over i starten af Excel: $RowsToSkip" -ForegroundColor Green
}
if ($MaxFilesToProcess -gt 0) {
    Write-Host "[+] Maks filer der behandles i koerslen: $MaxFilesToProcess" -ForegroundColor Green
} else {
    Write-Host "[+] Maks filer der behandles i koerslen: Alle" -ForegroundColor Green
}
Write-Host ""

# Auto-detect Excel file
$scriptDir = if ($PSScriptRoot) {
    $PSScriptRoot
} elseif ($MyInvocation.MyCommand.Path) {
    Split-Path -Parent $MyInvocation.MyCommand.Path
} else {
    Get-Location
}

if (-not $ExcelFile) {
    $excelFiles = Get-ChildItem -Path $scriptDir -Filter "*.xlsx" -File
    if ($excelFiles.Count -eq 0) {
        $excelFiles = Get-ChildItem -Path $scriptDir -Filter "*.xls" -File
    }

    if ($excelFiles.Count -eq 0) {
        Write-Host "FEJL: Ingen Excel fil (.xlsx eller .xls) fundet!" -ForegroundColor Red
        pause; exit 1
    } elseif ($excelFiles.Count -eq 1) {
        $ExcelFile = $excelFiles[0].FullName
        Write-Host "[+] Fundet Excel fil: $($excelFiles[0].Name)" -ForegroundColor Green
    } else {
        Write-Host "Fundet flere Excel filer:" -ForegroundColor Yellow
        for ($i = 0; $i -lt $excelFiles.Count; $i++) {
            Write-Host "  [$($i+1)] $($excelFiles[$i].Name)" -ForegroundColor Cyan
        }
        $selection = Read-Host "`nVaelg fil nummer (1-$($excelFiles.Count))"
        $selectedIndex = [int]$selection - 1
        if ($selectedIndex -ge 0 -and $selectedIndex -lt $excelFiles.Count) {
            $ExcelFile = $excelFiles[$selectedIndex].FullName
            Write-Host "[+] Valgt: $($excelFiles[$selectedIndex].Name)" -ForegroundColor Green
        } else {
            Write-Host "Ugyldig valg!" -ForegroundColor Red
            exit 1
        }
    }
}

Write-Host ""

# Prompt for credentials
if (-not $Username) { $Username = Read-Host "P360 Brugernavn (fx DOMAIN\brugernavn)" }
if (-not $Password) {
    $SecurePassword = Read-Host "P360 Password" -AsSecureString
    $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecurePassword)
    $Password = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
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

Write-Host ""
Write-Host "[+] Excel fil: $ExcelFile" -ForegroundColor Green
Write-Host "[+] Filtype: $FileType" -ForegroundColor Green
Write-Host "[+] Brugernavn: $Username" -ForegroundColor Green
Write-Host "[+] Output mappe: $OutputDir" -ForegroundColor Green
Write-Host ""

# Create output directories
if (-not (Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir | Out-Null
    Write-Host "[+] Oprettet output mappe: $OutputDir" -ForegroundColor Green
}

if (-not (Test-Path $MarkdownDir)) {
    New-Item -ItemType Directory -Path $MarkdownDir | Out-Null
    Write-Host "[+] Oprettet markdown mappe: $MarkdownDir" -ForegroundColor Green
}

# Read Excel file
if (-not $convertOnly) {
    Write-Host ""
    Write-Host "[*] Laeser Excel fil..." -ForegroundColor Yellow

try {
    $maxDataRowsToRead = if ($MaxFilesToProcess -gt 0) { $MaxFilesToProcess } else { 0 }

    # Try using ImportExcel module first (if available)
    if (Get-Module -ListAvailable -Name ImportExcel) {
        Import-Module ImportExcel
        $startRow = 1 + $RowsToSkip

        if ($maxDataRowsToRead -gt 0) {
            $endRow = $startRow + $maxDataRowsToRead
            $data = Import-Excel -Path $ExcelFile -StartRow $startRow -EndRow $endRow
            Write-Host "[+] Bruger ImportExcel module (StartRow=$startRow, EndRow=$endRow)" -ForegroundColor Green
        } else {
            $data = Import-Excel -Path $ExcelFile -StartRow $startRow
            Write-Host "[+] Bruger ImportExcel module (StartRow=$startRow)" -ForegroundColor Green
        }
    } else {
        # Fallback to Excel COM object
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        
        $workbook = $excel.Workbooks.Open($ExcelFile)
        $worksheet = $workbook.Worksheets.Item(1)
        
        # Get used range
        $usedRange = $worksheet.UsedRange
        $rowCount = $usedRange.Rows.Count
        $colCount = $usedRange.Columns.Count
        
        Write-Host "[+] Excel aabn: $rowCount raekker, $colCount kolonner" -ForegroundColor Green
        
        # Read headers
        $headerRow = 1 + $RowsToSkip
        if ($headerRow -gt $rowCount) {
            throw "Header row ($headerRow) er udenfor dataomraadet ($rowCount raekker)"
        }

        $headers = @{}
        for ($col = 1; $col -le $colCount; $col++) {
            $headerCell = $worksheet.Cells.Item($headerRow, $col)
            $headerName = $headerCell.Text
            if ($headerName) {
                # Normalize header names (fx 'DocID(D)(P)' med evt. linjeskift)
                $cleanHeader = $headerName.Trim() -replace '\s*\(D\)\(P\)\s*$', ''
                $headers[$cleanHeader] = $col
            }
        }
        
        Write-Host "[+] Fundet $($headers.Count) kolonner" -ForegroundColor Green
        
        # Build data array
        $data = @()
        $firstDataRow = $headerRow + 1
        $availableDataRows = [Math]::Max(0, $rowCount - $headerRow)
        $rowsToRead = if ($maxDataRowsToRead -gt 0) {
            [Math]::Min($maxDataRowsToRead, $availableDataRows)
        } else {
            $availableDataRows
        }

        if ($rowsToRead -gt 100) {
            Write-Host "[+] Behandler data i blokke af 100 rækker" -ForegroundColor DarkGray
        }

        $lastDataRow = if ($rowsToRead -gt 0) {
            $firstDataRow + $rowsToRead - 1
        } else {
            $firstDataRow - 1
        }

        for ($row = $firstDataRow; $row -le $lastDataRow; $row++) {
            $rowData = @{}
            foreach ($header in $headers.Keys) {
                $col = $headers[$header]
                $cellValue = $worksheet.Cells.Item($row, $col).Text
                $rowData[$header] = $cellValue
            }
            $data += [PSCustomObject]$rowData
            
            # Progress indicator
            if ($rowsToRead -gt 100) {
                $processedRows = ($row - $firstDataRow + 1)
                if (($processedRows % 100 -eq 0) -or ($row -eq $lastDataRow)) {
                    Write-Host "  Laeser raekke $processedRows / $rowsToRead..." -ForegroundColor Gray
                }
            }
        }
        
        # Cleanup
        $workbook.Close($false)
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        
        Write-Host "[+] Excel lukket" -ForegroundColor Green
    }
    
    Write-Host "[+] Indlaest $($data.Count) raekker" -ForegroundColor Green
    
} catch {
    Write-Host "FEJL: Kunne ikke laese Excel fil" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    pause
    exit 1
}
} # End if not convertOnly

if (-not $convertOnly) {
    Write-Host ""
    Write-Host "[*] Filtrerer raekker (KUN DOC/DOCX/PDF + Afgørelse-format)..." -ForegroundColor Yellow

# Filter data
$allowedExtensions = @('DOC', 'DOCX', 'PDF')
if ($FileType -eq 'word') { $allowedExtensions = @('DOC', 'DOCX') }
elseif ($FileType -eq 'pdf') { $allowedExtensions = @('PDF') }

$filesToDownload = @()
$skipped = 0
$skipReasons = @{}
$decisionTitlePattern = '^Afg.relse af'
$decisionFileNamePattern = 'Afg.relse'

foreach ($row in $data) {
    $docNameCandidateKeys = @(
        'DocName(D)(P)',
        'DocName',
        'DocumentTitle',
        'Title'
    )
    $fileNameCandidateKeys = @(
        'FileNameText(D)(P)',
        'FileNameText',
        'FileName',
        'ToFile.Filename'
    )

    # Extract fields
    $docName = ""
    foreach ($key in $docNameCandidateKeys) {
        if ($row.PSObject.Properties[$key] -and -not [string]::IsNullOrWhiteSpace([string]$row.$key)) {
            $docName = [string]$row.$key
            break
        }
    }

    $fileRecno = if ($row.FileRecno) { $row.FileRecno } else { "" }
    $fileName = ""
    foreach ($key in $fileNameCandidateKeys) {
        if ($row.PSObject.Properties[$key] -and -not [string]::IsNullOrWhiteSpace([string]$row.$key)) {
            $fileName = [string]$row.$key
            break
        }
    }
    $fileFormat = if ($row.'ToFile.Type') { $row.'ToFile.Type' }
                elseif ($row.'ToFile.Format') { $row.'ToFile.Format' }
                else { "" }
    $importedDocNo = if ($row.ImportedDocumentNumber) { $row.ImportedDocumentNumber } else { "" }
    $documentRecnoRaw = if ($row.recno) { $row.recno }
                        elseif ($row.DocID) { $row.DocID }
                        else { "" }
    $documentRecno = ""
    if ($documentRecnoRaw -match '(\d+)') {
        $documentRecno = $Matches[1]
    }
    $klassifikation = if ($row.'ToClassification.Code') { $row.'ToClassification.Code' } else { "" }
    $caseTitle = if ($row.CaseNameAndDescription) { $row.CaseNameAndDescription } else { "" }
    
    # Document-level validation
    $skip = $false
    $skipReason = ""
    
    # Rule 1: Document title must start with "Afgørelse af"
    if ($docName -notmatch $decisionTitlePattern) {
        $skip = $true
        $skipReason = "Titel starter ikke med Afgørelse af"
    }
    
    # Rule 2: Check klassifikation
    if (-not $skip -and $klassifikation -match '2100|EFTERLEVELSE') {
        $skip = $true
        $skipReason = "Klassifikation=Efterlevelse"
    }
    
    # Rule 3: Check case title
    if (-not $skip -and $caseTitle -match 'EFTERLEVELSE|OMKOSTNINGSDAEKNING') {
        $skip = $true
        $skipReason = "Sagstitel=Ekskluderet"
    }
    
    # Rule 4: Check document title for EFTERLEVELSE
    if (-not $skip -and $docName -match 'EFTERLEVELSE') {
        $skip = $true
        $skipReason = "Dokumenttitel=Efterlevelse"
    }
    
    # Normalize extension
    $extension = ""
    if ($fileFormat -match '(DOC|DOCX|PDF)') { 
        $extension = $Matches[1].ToUpper() 
    } elseif ($fileName -match '\.(?i)(pdf|docx?)$') { 
        $extension = $Matches[1].ToUpper() 
    }
    
    # Rule 5: Must have valid extension
    if (-not $skip -and $extension -notin $allowedExtensions) {
        $skip = $true
        $skipReason = "Extension=$extension"
    }

    # Rule 6: Filename must contain "Afgørelse"
    if (-not $skip -and $fileName -notmatch $decisionFileNamePattern) {
        $skip = $true
        $skipReason = "Filnavn mangler Afgørelse"
    }
    
    # Extract numeric FileID
    $fileId = ""
    if ($fileRecno -match '^\d+$') {
        $fileId = $fileRecno
    } elseif ($fileRecno -match '(\d+)') {
        $fileId = $Matches[1]
    }
    
    # Rule 7: Must have numeric FileID
    if (-not $skip -and -not ($fileId -match '^\d+$')) {
        $skip = $true
        $skipReason = "FileID=ugyldig"
    }
    
    if ($skip) {
        $skipped++
        if (-not $skipReasons.ContainsKey($skipReason)) {
            $skipReasons[$skipReason] = 0
        }
        $skipReasons[$skipReason]++
        
        if ($skipped -le 10) {
            Write-Host "o SKIP: $skipReason | Titel='$docName'" -ForegroundColor DarkGray
        }
    } else {
        # Extract case number from document number (e.g., "20/01453-6" -> "20/01453")
        $caseNumber = ""
        if ($importedDocNo -match '^(\d{2}/\d{5})') {
            $caseNumber = $Matches[1]
        }
        
        # Build filename: "FileID XX_YYYYY Afgørelse.ext"
        $safeCaseNumber = if ($caseNumber) { $caseNumber -replace '/', '_' } else { "UkendtSagsnummer" }
        $newFilename = "$fileId $safeCaseNumber Afgørelse.$($extension.ToLower())"
        
        Write-Host "v OK: '$docName' | Ext='$extension' | FileId='$fileId'" -ForegroundColor Green
        
        $filesToDownload += @{
            FileId = $fileId
            Filename = $newFilename
            DocumentTitle = $docName
            Extension = $extension
            DocumentNumber = $importedDocNo
            CaseTitle = $caseTitle
            CaseNumber = $caseNumber
            DocumentRecno = $documentRecno
        }
    }
}

Write-Host ""
if ($MaxFilesToProcess -gt 0 -and $filesToDownload.Count -gt $MaxFilesToProcess) {
    $filesToDownload = @($filesToDownload | Select-Object -First $MaxFilesToProcess)
    Write-Host "[!] Begraenser til foerste $MaxFilesToProcess filer" -ForegroundColor Yellow
}

Write-Host "[+] $($filesToDownload.Count) filer klar til download" -ForegroundColor Green
Write-Host "[+] $skipped filer sprunget over:" -ForegroundColor Yellow
foreach ($reason in $skipReasons.Keys | Sort-Object) {
    Write-Host "    - $reason`: $($skipReasons[$reason])" -ForegroundColor Gray
}
Write-Host ""

if ($filesToDownload.Count -eq 0) {
    Write-Host "INGEN filer at hente!" -ForegroundColor Red
    pause; exit 0
}

$confirm = Read-Host "Start download? (Y/N)"
if ($confirm -ne 'Y' -and $confirm -ne 'y') {
    Write-Host "Afbrudt" -ForegroundColor Yellow
    exit 0
}
} # End if not convertOnly

# Resolve pdftotext path once for all conversion flows
$pdfToTextPath = Resolve-PdfToTextPath -ScriptDir $scriptDir

# Download or scan existing files
$downloaded = 0
$skippedExisting = 0

if (-not $convertOnly) {
    # DOWNLOAD MODE
    # Prepare credentials for download
$baseUrl = 'https://esdh-nh-arkiv'
$securePass = ConvertTo-SecureString $Password -AsPlainText -Force
$credential = New-Object System.Management.Automation.PSCredential($Username, $securePass)

Write-Host ""
Write-Host "====================================================================" -ForegroundColor Cyan
Write-Host " STARTER DOWNLOAD" -ForegroundColor Cyan
Write-Host "====================================================================" -ForegroundColor Cyan
Write-Host ""

$downloaded = 0
$errors = 0
$skippedExisting = 0
$downloadedFiles = @()

foreach ($file in $filesToDownload) {
    $fileNum = $downloaded + $errors + $skippedExisting + 1
    $total = $filesToDownload.Count
    $url = "$baseUrl/GetFile.aspx?fileId=$($file.FileId)&redirect=true"
    
    # Sanitize filename (replace / with _)
    $safeFilename = $file.Filename -replace '[\\/:*?"<>|]', '_'
    $outputPath = Join-Path $OutputDir $safeFilename
    
    # Check if file already exists
    if (Test-Path $outputPath) {
        Write-Host "- [$fileNum/$total] '$($file.DocumentTitle)' -> FINDES ALLEREDE (springer over)" -ForegroundColor Yellow
        $skippedExisting++
        
        # Add to downloaded list anyway for markdown
        $downloadedFiles += @{
            Path = $outputPath
            Title = $file.DocumentTitle
            Extension = $file.Extension
            FileId = $file.FileId
            DocumentNumber = $file.DocumentNumber
            CaseNumber = $file.CaseNumber
            DocumentRecno = $file.DocumentRecno
            Filename = $safeFilename
        }

        Write-Host "    [*] Konverterer straks til markdown..." -ForegroundColor DarkGray
        $null = Convert-DownloadedFileToMarkdown -FileInfo $downloadedFiles[-1] -MarkdownDir $MarkdownDir -PdfToTextPath $pdfToTextPath
        continue
    }
    
    Write-Host "v [$fileNum/$total] '$($file.DocumentTitle)' -> Henter..." -ForegroundColor Cyan
    
    try {
        $startTime = Get-Date
        Invoke-WebRequest -Uri $url -Credential $credential -OutFile $outputPath -UseBasicParsing
        $duration = (Get-Date) - $startTime
        $fileSize = (Get-Item $outputPath).Length / 1MB
        Write-Host "v [$fileNum/$total] '$($file.DocumentTitle)' -> OK ($("{0:N2}" -f $fileSize) MB, $("{0:N1}" -f $duration.TotalSeconds)s)" -ForegroundColor Green
        $downloaded++
        
        $downloadedFiles += @{
            Path = $outputPath
            Title = $file.DocumentTitle
            Extension = $file.Extension
            FileId = $file.FileId
            DocumentNumber = $file.DocumentNumber
            CaseNumber = $file.CaseNumber
            DocumentRecno = $file.DocumentRecno
            Filename = $safeFilename
        }

        Write-Host "    [*] Konverterer straks til markdown..." -ForegroundColor DarkGray
        $null = Convert-DownloadedFileToMarkdown -FileInfo $downloadedFiles[-1] -MarkdownDir $MarkdownDir -PdfToTextPath $pdfToTextPath
    } catch {
        Write-Host "x [$fileNum/$total] '$($file.DocumentTitle)' -> FEJL ($($_.Exception.Message))" -ForegroundColor Red
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
Write-Host "- $skippedExisting filer eksisterede allerede" -ForegroundColor Yellow
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
            if ($basename -match '^(\d{2}[_/]\d{5})') {
                $caseNumber = $Matches[1] -replace '_', '/'
            }
            
            $downloadedFiles += @{
                Path = $file.FullName
                Title = $basename
                Extension = $extension
                FileId = ""
                DocumentNumber = ""
                CaseNumber = $caseNumber
                DocumentRecno = ""
                Filename = $file.Name
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

# Convert files to markdown and create index
if ($downloadedFiles.Count -gt 0) {
    Write-Host ""

    # In download mode we convert each file immediately like SIF flow.
    # In convert-only mode we convert all existing files one by one here.
    if ($convertOnly) {
        Write-Host "====================================================================" -ForegroundColor Cyan
        Write-Host " KONVERTERER TIL MARKDOWN" -ForegroundColor Cyan
        Write-Host "====================================================================" -ForegroundColor Cyan
        Write-Host ""

        $converted = 0
        $conversionErrors = 0

        foreach ($file in $downloadedFiles) {
            $fileNum = $converted + $conversionErrors + 1
            $total = $downloadedFiles.Count
            Write-Host "  [$fileNum/$total] $($file.Filename)" -ForegroundColor Cyan

            try {
                $null = Convert-DownloadedFileToMarkdown -FileInfo $file -MarkdownDir $MarkdownDir -PdfToTextPath $pdfToTextPath
                $converted++
            } catch {
                Write-Host "    FEJL: $($_.Exception.Message)" -ForegroundColor Red
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

    # Create INDEX.md
    Write-Host "[*] Opretter markdown index..." -ForegroundColor Yellow
    
    $indexContent = @"
# P360 Dokumenter - Arkiv (Excel)

**Hentet:** $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')  
**Excel fil:** $ExcelFile  
**Antal filer:** $($downloadedFiles.Count)  
**Nye downloads:** $downloaded  
**Eksisterende:** $skippedExisting  

## Dokumenter

"@

    foreach ($file in $downloadedFiles) {
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($file.Filename)
        $markdownFile = "$baseName.md"
        
        $indexContent += @"

### $($file.Title)

- **Dokument nummer:** $($file.DocumentNumber)
- **Sags nummer:** $($file.CaseNumber)
- **Fil:** ``$($file.Filename)`` ($($file.Extension))
- **FileID:** $($file.FileId)
- **Lokal fil:** [``$($file.Filename)``](../arkiv_downloads/$($file.Filename))
- **Markdown:** [``$markdownFile``](./$markdownFile)

"@
    }
    
    $indexPath = Join-Path $MarkdownDir "INDEX.md"
    $utf8 = New-Object System.Text.UTF8Encoding $true
    [System.IO.File]::WriteAllText($indexPath, $indexContent, $utf8)
    
    Write-Host "[+] Markdown index oprettet: $indexPath" -ForegroundColor Green
}

Write-Host ""
Write-Host "Filer gemt i:" -ForegroundColor Cyan
Write-Host "  Downloads: $OutputDir" -ForegroundColor Cyan
if ($downloadedFiles.Count -gt 0) {
    Write-Host "  Markdown: $MarkdownDir" -ForegroundColor Cyan
}
Write-Host ""

$openFolder = Read-Host "Aaben download mappe? (Y/N)"
if ($openFolder -eq 'Y' -or $openFolder -eq 'y') {
    Invoke-Item $OutputDir
}

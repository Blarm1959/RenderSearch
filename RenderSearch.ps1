[CmdletBinding()]
param(
    [string]$Folder,
    [string[]]$Terms = @(
        'k rend','k-rend','krend','weber','parex','silver pearl','pearl',
        'render','rendering','rendered','monocouche','silicone','through colour','through color',
        'off-white','off white','colour','color','finish','ral',
        'white','cream','ivory','limestone','chalk','buff','beige'
    ),
    [string]$OutputFile,
    [int]$ContextCharacters = 220,
    [string[]]$PriorityTerms = @('k rend','k-rend','krend','weber','parex','silver pearl','pearl')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Write-Status {
    param([string]$Text, [ConsoleColor]$Colour = [ConsoleColor]::Gray)
    Write-Host $Text -ForegroundColor $Colour
}

function Resolve-DefaultSearchFolder {
    $oneDrive = $env:OneDrive
    if ([string]::IsNullOrWhiteSpace($oneDrive)) {
        $oneDrive = $env:OneDriveConsumer
    }
    if ([string]::IsNullOrWhiteSpace($oneDrive)) {
        throw 'OneDrive could not be located automatically. Run again with -Folder "C:\path\to\WDLBlandings\ReBuild".'
    }
    return (Join-Path $oneDrive 'WDLBlandings\ReBuild')
}

function Get-ZipXmlText {
    param([string]$Path, [ValidateSet('docx','xlsx')] [string]$Type)

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive = [System.IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $parts = if ($Type -eq 'docx') {
            $archive.Entries | Where-Object { $_.FullName -match '^word/(document|header\d*|footer\d*|footnotes|endnotes)\.xml$' }
        } else {
            $archive.Entries | Where-Object { $_.FullName -match '^xl/(sharedStrings|worksheets/sheet\d+)\.xml$' }
        }

        $builder = [System.Text.StringBuilder]::new()
        foreach ($entry in $parts) {
            $reader = [System.IO.StreamReader]::new($entry.Open())
            try {
                $xmlText = $reader.ReadToEnd()
                $plain = [System.Net.WebUtility]::HtmlDecode(($xmlText -replace '<[^>]+>', ' '))
                [void]$builder.AppendLine($plain)
            }
            finally { $reader.Dispose() }
        }
        return $builder.ToString()
    }
    finally { $archive.Dispose() }
}

function Get-OfficeComText {
    param([string]$Path, [ValidateSet('word','excel')] [string]$Application)

    if ($Application -eq 'word') {
        $word = $null; $doc = $null
        try {
            $word = New-Object -ComObject Word.Application
            $word.Visible = $false
            $word.DisplayAlerts = 0
            $doc = $word.Documents.Open($Path, $false, $true)
            return $doc.Content.Text
        }
        finally {
            if ($null -ne $doc) { $doc.Close($false) | Out-Null; [void][Runtime.InteropServices.Marshal]::ReleaseComObject($doc) }
            if ($null -ne $word) { $word.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word) }
            [GC]::Collect(); [GC]::WaitForPendingFinalizers()
        }
    }

    $excel = $null; $book = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $book = $excel.Workbooks.Open($Path, 0, $true)
        $builder = [System.Text.StringBuilder]::new()
        foreach ($sheet in $book.Worksheets) {
            $range = $sheet.UsedRange
            try {
                $values = $range.Value2
                if ($values -is [Array]) {
                    foreach ($value in $values) { if ($null -ne $value) { [void]$builder.AppendLine([string]$value) } }
                } elseif ($null -ne $values) {
                    [void]$builder.AppendLine([string]$values)
                }
            }
            finally { if ($null -ne $range) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($range) } }
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($sheet)
        }
        return $builder.ToString()
    }
    finally {
        if ($null -ne $book) { $book.Close($false) | Out-Null; [void][Runtime.InteropServices.Marshal]::ReleaseComObject($book) }
        if ($null -ne $excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
}

function Get-PdfText {
    param([string]$Path)

    $pdfToText = Get-Command pdftotext.exe -ErrorAction SilentlyContinue
    if ($null -eq $pdfToText) {
        throw 'PDF support requires pdftotext.exe. Word PDF conversion is deliberately not used because it can hang on individual PDFs.'
    }

    $temp = [IO.Path]::GetTempFileName()
    try {
        & $pdfToText.Source -layout -enc UTF-8 -- $Path $temp 2>$null
        if ($LASTEXITCODE -ne 0) {
            throw "pdftotext.exe returned exit code $LASTEXITCODE"
        }
        if (Test-Path $temp) { return [IO.File]::ReadAllText($temp) }
        return ''
    }
    finally { Remove-Item $temp -Force -ErrorAction SilentlyContinue }
}

function Get-FileText {
    param([IO.FileInfo]$File)

    switch ($File.Extension.ToLowerInvariant()) {
        '.txt'  { return [IO.File]::ReadAllText($File.FullName) }
        '.csv'  { return [IO.File]::ReadAllText($File.FullName) }
        '.log'  { return [IO.File]::ReadAllText($File.FullName) }
        '.md'   { return [IO.File]::ReadAllText($File.FullName) }
        '.xml'  { return [IO.File]::ReadAllText($File.FullName) }
        '.json' { return [IO.File]::ReadAllText($File.FullName) }
        '.docx' { return Get-ZipXmlText -Path $File.FullName -Type docx }
        '.xlsx' { return Get-ZipXmlText -Path $File.FullName -Type xlsx }
        '.doc'  { return Get-OfficeComText -Path $File.FullName -Application word }
        '.xls'  { return Get-OfficeComText -Path $File.FullName -Application excel }
        '.pdf'  { return Get-PdfText -Path $File.FullName }
        default { return $null }
    }
}

function Get-MatchesWithContext {
    param([string]$Text, [string[]]$SearchTerms, [int]$Context)

    if ([string]::IsNullOrWhiteSpace($Text)) { return @() }
    $normalised = $Text -replace '\s+', ' '
    $found = [System.Collections.Generic.List[object]]::new()

    foreach ($term in $SearchTerms) {
        $startAt = 0
        while ($startAt -lt $normalised.Length) {
            $index = $normalised.IndexOf($term, $startAt, [StringComparison]::OrdinalIgnoreCase)
            if ($index -lt 0) { break }
            $from = [Math]::Max(0, $index - $Context)
            $to = [Math]::Min($normalised.Length, $index + $term.Length + $Context)
            $snippet = $normalised.Substring($from, $to - $from).Trim()
            $found.Add([pscustomobject]@{ Term = $term; Context = $snippet })
            $startAt = $index + [Math]::Max(1, $term.Length)
        }
    }
    return $found
}


function Get-SkipReasonCategory {
    param([string]$Message)

    if ($Message -like 'PDF support requires pdftotext.exe*') { return 'PDF extractor unavailable' }
    if ($Message -like '*pdftotext.exe returned exit code*') { return 'PDF extraction failed' }
    if ($Message -like '*COM*' -or $Message -like '*Word*' -or $Message -like '*Excel*') { return 'Office extraction failed' }
    return 'Unreadable/extraction error'
}

if ([string]::IsNullOrWhiteSpace($Folder)) { $Folder = Resolve-DefaultSearchFolder }
$Folder = [IO.Path]::GetFullPath($Folder)
if (-not (Test-Path -LiteralPath $Folder -PathType Container)) { throw "Search folder does not exist: $Folder" }

if ([string]::IsNullOrWhiteSpace($OutputFile)) {
    $OutputFile = Join-Path $PSScriptRoot ("RenderSearch-Results-{0}.txt" -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
}

$extensions = @('.pdf','.doc','.docx','.xls','.xlsx','.txt','.csv','.log','.md','.xml','.json')
$files = @(Get-ChildItem -LiteralPath $Folder -File -Recurse -ErrorAction SilentlyContinue | Where-Object {
    $relativePath = $_.FullName.Substring($Folder.Length).TrimStart('\')
    $pathParts = $relativePath -split '[\\/]'
    ('VATReclaim' -notin $pathParts) -and ($extensions -contains $_.Extension.ToLowerInvariant())
})
$results = [System.Collections.Generic.List[object]]::new()
$errors = [System.Collections.Generic.List[object]]::new()

Write-Status ''
Write-Status '=========================================================' Cyan
Write-Status ' RenderSearch - Local File Search' Cyan
Write-Status '=========================================================' Cyan
Write-Status "Folder : $Folder"
Write-Status "Files  : $($files.Count)"
Write-Status "Terms  : $($Terms -join ', ')"
Write-Status ''

$pdfToTextAvailable = $null -ne (Get-Command pdftotext.exe -ErrorAction SilentlyContinue)
if (-not $pdfToTextAvailable -and ($files | Where-Object { $_.Extension -ieq '.pdf' })) {
    Write-Status '[WARN] pdftotext.exe is not installed. PDFs will be skipped rather than opened through Word.' DarkYellow
    Write-Status ''
}

$number = 0
foreach ($file in $files) {
    $number++
    Write-Progress -Activity 'Searching files' -Status $file.Name -PercentComplete (($number / [Math]::Max(1,$files.Count)) * 100)
    Write-Status ("[SCAN {0}/{1}] {2}" -f $number, $files.Count, $file.FullName) DarkGray
    try {
        $text = Get-FileText -File $file
        $matches = @(Get-MatchesWithContext -Text $text -SearchTerms $Terms -Context $ContextCharacters)
        foreach ($match in $matches) {
            $results.Add([pscustomobject]@{
                File = $file.Name
                Path = $file.FullName
                Type = $file.Extension.ToLowerInvariant()
                Term = $match.Term
                Context = $match.Context
            })
        }
        if ($matches.Count -gt 0) { Write-Status ("[FOUND] {0} ({1} match{2})" -f $file.FullName, $matches.Count, $(if($matches.Count -eq 1){''}else{'es'})) Green }
    }
    catch {
        $errors.Add([pscustomobject]@{ Path = $file.FullName; Type = $file.Extension.ToLowerInvariant(); Reason = (Get-SkipReasonCategory -Message $_.Exception.Message); Error = $_.Exception.Message })
        Write-Status "[SKIP] $($file.FullName) - $($_.Exception.Message)" DarkYellow
    }
}
Write-Progress -Activity 'Searching files' -Completed

$report = [System.Text.StringBuilder]::new()
[void]$report.AppendLine('RenderSearch Results')
[void]$report.AppendLine(('Generated: {0}' -f (Get-Date)))
[void]$report.AppendLine(('Folder:    {0}' -f $Folder))
[void]$report.AppendLine(('Files:     {0}' -f $files.Count))
[void]$report.AppendLine(('Matches:   {0}' -f $results.Count))
[void]$report.AppendLine(('Skipped:   {0}' -f $errors.Count))
[void]$report.AppendLine('')


$priorityResults = @($results | Where-Object {
    $resultTerm = $_.Term
    $PriorityTerms | Where-Object { $resultTerm -ieq $_ } | Select-Object -First 1
})

if ($priorityResults.Count -gt 0) {
    [void]$report.AppendLine(('=' * 90))
    [void]$report.AppendLine('HIGH PRIORITY MANUFACTURER MATCHES')
    [void]$report.AppendLine(('=' * 90))
    foreach ($group in ($priorityResults | Group-Object Path)) {
        [void]$report.AppendLine($group.Name)
        foreach ($item in $group.Group) {
            [void]$report.AppendLine(('TERM: {0}' -f $item.Term))
            [void]$report.AppendLine($item.Context)
            [void]$report.AppendLine('')
        }
    }
}

foreach ($group in ($results | Group-Object Path)) {
    [void]$report.AppendLine(('=' * 90))
    [void]$report.AppendLine($group.Name)
    [void]$report.AppendLine(('=' * 90))
    foreach ($item in $group.Group) {
        [void]$report.AppendLine(('TERM: {0}' -f $item.Term))
        [void]$report.AppendLine($item.Context)
        [void]$report.AppendLine('')
    }
}

if ($errors.Count -gt 0) {
    [void]$report.AppendLine(('=' * 90))
    [void]$report.AppendLine('SKIPPED FILE SUMMARY')
    [void]$report.AppendLine(('=' * 90))
    [void]$report.AppendLine('By file type:')
    foreach ($group in ($errors | Group-Object Type | Sort-Object @{Expression='Count'; Descending=$true}, @{Expression='Name'; Descending=$false})) {
        [void]$report.AppendLine(('  {0,-8} {1,5}' -f $group.Name, $group.Count))
    }
    [void]$report.AppendLine('')
    [void]$report.AppendLine('By reason:')
    foreach ($group in ($errors | Group-Object Reason | Sort-Object @{Expression='Count'; Descending=$true}, @{Expression='Name'; Descending=$false})) {
        [void]$report.AppendLine(('  {0,-32} {1,5}' -f $group.Name, $group.Count))
    }
    [void]$report.AppendLine('')

    [void]$report.AppendLine(('=' * 90))
    [void]$report.AppendLine('FILES THAT COULD NOT BE READ')
    [void]$report.AppendLine(('=' * 90))
    foreach ($item in $errors) { [void]$report.AppendLine("$($item.Path)`r`n  $($item.Error)`r`n") }
}

[IO.File]::WriteAllText($OutputFile, $report.ToString(), [Text.UTF8Encoding]::new($false))

Write-Status ''
Write-Status "Search complete: $($results.Count) matches in $((($results | Select-Object -ExpandProperty Path -Unique) | Measure-Object).Count) files." Cyan
Write-Status "Priority manufacturer matches: $($priorityResults.Count)"
Write-Status "Skipped/unreadable files: $($errors.Count)"
if ($errors.Count -gt 0) {
    Write-Status 'Skipped by file type:' DarkYellow
    foreach ($group in ($errors | Group-Object Type | Sort-Object @{Expression='Count'; Descending=$true}, @{Expression='Name'; Descending=$false})) {
        Write-Status ("  {0,-8} {1,5}" -f $group.Name, $group.Count) DarkYellow
    }
    Write-Status 'Skipped by reason:' DarkYellow
    foreach ($group in ($errors | Group-Object Reason | Sort-Object @{Expression='Count'; Descending=$true}, @{Expression='Name'; Descending=$false})) {
        Write-Status ("  {0,-32} {1,5}" -f $group.Name, $group.Count) DarkYellow
    }
}
Write-Status "Results: $OutputFile" Cyan
Write-Status ''

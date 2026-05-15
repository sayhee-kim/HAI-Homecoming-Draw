$ErrorActionPreference = 'Stop'

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$Port = 5173
$ExcelFile = (Get-ChildItem -LiteralPath $Root -Filter '*.xlsx' |
    Where-Object { $_.Name -notlike '~$*' } |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1).FullName

Add-Type -AssemblyName System.IO.Compression.FileSystem

function Convert-ColumnNameToNumber {
    param([string]$CellRef)
    $letters = ($CellRef -replace '\d', '').ToUpperInvariant()
    $number = 0
    foreach ($char in $letters.ToCharArray()) {
        $number = ($number * 26) + ([int][char]$char - [int][char]'A' + 1)
    }
    return $number
}

function Get-SharedStrings {
    param([string]$Folder)
    $path = Join-Path $Folder 'xl\sharedStrings.xml'
    if (-not (Test-Path $path)) { return @() }

    [xml]$xml = Get-Content -LiteralPath $path -Raw -Encoding UTF8
    $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
    $ns.AddNamespace('x', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')

    $strings = New-Object System.Collections.Generic.List[string]
    foreach ($si in $xml.SelectNodes('//x:si', $ns)) {
        $parts = New-Object System.Collections.Generic.List[string]
        foreach ($textNode in $si.SelectNodes('.//x:t', $ns)) {
            $parts.Add($textNode.InnerText)
        }
        $strings.Add(($parts -join ''))
    }
    return $strings.ToArray()
}

function Get-CellValue {
    param($Cell, [string[]]$SharedStrings)
    $valueNode = $Cell.SelectSingleNode('*[local-name()="v"]')
    if ($null -eq $valueNode) { return '' }

    $raw = $valueNode.InnerText
    if ($Cell.t -eq 's') {
        $index = 0
        if ([int]::TryParse($raw, [ref]$index) -and $index -lt $SharedStrings.Count) {
            return $SharedStrings[$index].Trim()
        }
    }
    return $raw.Trim()
}

function Read-Participants {
    if (-not $ExcelFile -or -not (Test-Path $ExcelFile)) {
        throw "Cannot find an .xlsx participant file in this folder."
    }

    $temp = Join-Path $env:TEMP ("teachersday_roulette_" + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $temp | Out-Null
    try {
        $copy = Join-Path $temp 'book.zip'
        Copy-Item -LiteralPath $ExcelFile -Destination $copy
        [System.IO.Compression.ZipFile]::ExtractToDirectory($copy, (Join-Path $temp 'book'))

        $book = Join-Path $temp 'book'
        $shared = Get-SharedStrings -Folder $book
        [xml]$sheet = Get-Content -LiteralPath (Join-Path $book 'xl\worksheets\sheet1.xml') -Raw -Encoding UTF8
        $ns = New-Object System.Xml.XmlNamespaceManager($sheet.NameTable)
        $ns.AddNamespace('x', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')

        $people = New-Object System.Collections.Generic.List[object]
        foreach ($row in $sheet.SelectNodes('//x:sheetData/x:row', $ns)) {
            $cells = @{}
            foreach ($cell in $row.SelectNodes('x:c', $ns)) {
                $col = Convert-ColumnNameToNumber $cell.r
                $cells[$col] = Get-CellValue -Cell $cell -SharedStrings $shared
            }
            $rowNumber = [int]$row.r
            $leftName = $cells.Item(2)
            $leftAttendance = $cells.Item(3)
            $leftGroup = $cells.Item(1)
            $rightName = $cells.Item(8)
            $rightAttendance = $cells.Item(9)
            $sources = @(
                [pscustomobject]@{
                    CandidateName = $leftName
                    Attendance = $leftAttendance
                    Group = $leftGroup
                    Kind = 'student'
                },
                [pscustomobject]@{
                    CandidateName = $rightName
                    Attendance = $rightAttendance
                    Group = ''
                    Kind = 'alumni'
                }
            )

            foreach ($source in $sources) {
                $nameValue = $source.CandidateName
                $attendanceValue = $source.Attendance
                $groupValue = $source.Group
                $typeValue = $source.Kind
                $name = if ($nameValue) { $nameValue.Trim() } else { '' }
                $attendance = if ($attendanceValue) { $attendanceValue.Trim() } else { '' }
                if (
                    $rowNumber -ge 3 -and
                    $name.Length -gt 0 -and
                    $attendance -match '^(O|o|Y|Yes|YES|1)$'
                ) {
                    $people.Add([pscustomobject]@{
                        name = $name
                        group = if ($groupValue) { $groupValue.Trim() } else { $typeValue }
                        row = $rowNumber
                        type = $typeValue
                    })
                }
            }
        }

        return @($people.ToArray())
    }
    finally {
        if (Test-Path $temp) {
            Remove-Item -LiteralPath $temp -Recurse -Force
        }
    }
}

function Send-Response {
    param($Context, [int]$Status, [string]$ContentType, [byte[]]$Bytes)
    $Context.Response.StatusCode = $Status
    $Context.Response.ContentType = $ContentType
    $Context.Response.ContentLength64 = $Bytes.Length
    $Context.Response.OutputStream.Write($Bytes, 0, $Bytes.Length)
    $Context.Response.OutputStream.Close()
}

function Send-Json {
    param($Context, $Data, [int]$Status = 200)
    $json = $Data | ConvertTo-Json -Depth 5
    Send-Response $Context $Status 'application/json; charset=utf-8' ([Text.Encoding]::UTF8.GetBytes($json))
}

function Get-ContentType {
    param([string]$Path)
    switch ([IO.Path]::GetExtension($Path).ToLowerInvariant()) {
        '.html' { 'text/html; charset=utf-8' }
        '.css' { 'text/css; charset=utf-8' }
        '.js' { 'text/javascript; charset=utf-8' }
        '.png' { 'image/png' }
        '.jpg' { 'image/jpeg' }
        '.jpeg' { 'image/jpeg' }
        '.mp4' { 'video/mp4' }
        default { 'application/octet-stream' }
    }
}

$listener = [System.Net.HttpListener]::new()
$listener.Prefixes.Add("http://localhost:$Port/")
$listener.Start()

Write-Host ""
Write-Host "Teachers' Day Roulette is running."
Write-Host "Open: http://localhost:$Port"
Write-Host "Press Ctrl+C to stop."
Write-Host ""
Start-Process "http://localhost:$Port"

try {
    while ($listener.IsListening) {
        $context = $listener.GetContext()
        $path = [Uri]::UnescapeDataString($context.Request.Url.AbsolutePath)

        try {
            if ($path -eq '/api/participants') {
                $people = Read-Participants
                Send-Json $context @{
                    participants = $people
                    count = $people.Count
                    source = [IO.Path]::GetFileName($ExcelFile)
                    updatedAt = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
                }
                continue
            }

            if ($path -eq '/') { $path = '/index.html' }
            $relative = $path.TrimStart('/') -replace '/', [IO.Path]::DirectorySeparatorChar
            $file = Join-Path $Root $relative
            $resolvedRoot = (Resolve-Path $Root).Path
            $resolvedFile = if (Test-Path $file) { (Resolve-Path $file).Path } else { '' }

            if (-not $resolvedFile.StartsWith($resolvedRoot)) {
                Send-Json $context @{ error = 'Invalid path' } 403
                continue
            }

            if (Test-Path $resolvedFile -PathType Leaf) {
                $bytes = [IO.File]::ReadAllBytes($resolvedFile)
                Send-Response $context 200 (Get-ContentType $resolvedFile) $bytes
            }
            else {
                Send-Json $context @{ error = 'Not found' } 404
            }
        }
        catch {
            Send-Json $context @{ error = $_.Exception.Message } 500
        }
    }
}
finally {
    $listener.Stop()
}

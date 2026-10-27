# Read-only CorelDRAW environment and ordinary-group inventory. Windows PowerShell 5.1.
[CmdletBinding()]
param(
    [string]$ProgId = 'CorelDRAW.Application.18',
    [switch]$InspectRunning,
    [string]$DocumentFullName,
    [switch]$Inventory,
    [ValidateRange(1, 50)][int]$MaxDepth = 30,
    [string]$OutputPath
)
$ErrorActionPreference = 'Stop'
if ($Inventory -and [string]::IsNullOrWhiteSpace($DocumentFullName)) {
    throw '-Inventory requires -DocumentFullName; the active document is never assumed.'
}
if ($DocumentFullName -and -not $InspectRunning) {
    throw '-DocumentFullName requires -InspectRunning.'
}
if ($OutputPath) {
    $OutputPath = [IO.Path]::GetFullPath($OutputPath)
    if (Test-Path -LiteralPath $OutputPath) { throw 'OutputPath already exists; choose a new path.' }
    if (-not (Test-Path -LiteralPath ([IO.Path]::GetDirectoryName($OutputPath)) -PathType Container)) {
        throw 'OutputPath parent directory does not exist.'
    }
}
$issues = New-Object 'System.Collections.Generic.List[object]'
$objects = New-Object 'System.Collections.Generic.List[object]'
$documentRows = New-Object 'System.Collections.Generic.List[object]'
$pageRows = New-Object 'System.Collections.Generic.List[object]'
$ordinaryGroupsComplete = $true
$exitStatus = 0

function Read-Property {
    param($Target, [string]$Name, [string]$Context)
    try {
        if ($null -eq $Target) { throw 'Parent object unavailable.' }
        $value = $Target.$Name
        # PowerShell may suppress a getter exception and return null instead.
        # Treat null as unknown, never as zero overflow or valid geometry.
        if ($null -eq $value) { throw 'Property returned null or could not be read.' }
        return $value
    } catch {
        $issues.Add([pscustomobject]@{context=$Context; property=$Name; error=$_.Exception.Message})
        return $null
    }
}

function Read-Shapes {
    param($Shapes, [int]$PageIndex, [string]$ParentPath, [int]$Depth)
    if ($Depth -gt $MaxDepth) {
        $script:ordinaryGroupsComplete = $false
        $issues.Add([pscustomobject]@{context=$ParentPath; error='MaxDepth reached; subtree not read.'})
        return
    }
    try { $count = [int]$Shapes.Count } catch {
        $script:ordinaryGroupsComplete = $false
        $issues.Add([pscustomobject]@{context=$ParentPath; error=$_.Exception.Message})
        return
    }
    for ($shapeIndex=1; $shapeIndex -le $count; $shapeIndex++) {
        $objectPath = $ParentPath + '/' + $shapeIndex
        try { $shape = $Shapes.Item($shapeIndex) } catch {
            $script:ordinaryGroupsComplete = $false
            $issues.Add([pscustomobject]@{context=$objectPath; error=$_.Exception.Message})
            continue
        }
        $row = [ordered]@{page=$PageIndex; objectPath=$objectPath; depth=$Depth}
        foreach ($property in @('StaticID','Type','Name','LeftX','TopY','BottomY','SizeWidth','SizeHeight')) {
            $row[$property] = Read-Property $shape $property $objectPath
        }
        $layer = Read-Property $shape 'Layer' $objectPath
        $row['layerName'] = Read-Property $layer 'Name' $objectPath
        if ($null -eq $row['Type']) { $script:ordinaryGroupsComplete = $false }
        if ($row['Type'] -eq 6) {
            $text = Read-Property $shape 'Text' $objectPath
            $story = Read-Property $text 'Story' $objectPath
            $row['text'] = Read-Property $story 'Text' $objectPath
            $row['font'] = Read-Property $story 'Font' $objectPath
            $row['fontSizePt'] = Read-Property $story 'Size' $objectPath
            $row['textOverflow'] = Read-Property $text 'Overflow' $objectPath
        }
        $objects.Add([pscustomobject]$row)
        if ($row['Type'] -eq 7) {
            try { Read-Shapes $shape.Shapes $PageIndex $objectPath ($Depth+1) } catch {
                $script:ordinaryGroupsComplete = $false
                $issues.Add([pscustomobject]@{context=$objectPath; error=$_.Exception.Message})
            }
        }
    }
}

$registrations = New-Object 'System.Collections.Generic.List[object]'
foreach ($candidate in @($ProgId,'CorelDRAW.Application','CorelDRAW.Application.18') | Select-Object -Unique) {
    $key = $null
    try {
        $key = [Microsoft.Win32.Registry]::ClassesRoot.OpenSubKey($candidate + '\CLSID')
        if ($key) { $registrations.Add([pscustomobject]@{progId=$candidate; clsid=$key.GetValue('')}) }
    } finally { if ($key) { $key.Dispose() } }
}
$report = [ordered]@{
    generatedAt = [DateTimeOffset]::Now.ToString('o')
    readOnly = $true
    hostVersion = $PSVersionTable.PSVersion.ToString()
    is64BitProcess = [Environment]::Is64BitProcess
    apartmentState = [Threading.Thread]::CurrentThread.ApartmentState.ToString()
    requestedProgId = $ProgId
    checkedRegistrations = @($registrations.ToArray())
    registrationScope = 'Requested ProgID, version-independent ProgID, historical X8 ProgID; not all installed versions.'
    runningProcesses = @(Get-Process -Name 'CorelDRW*' -ErrorAction SilentlyContinue | Select-Object Id,ProcessName)
    connection = 'not_requested'
    applicationVersion = $null
    documents = @()
    target = $null
    pages = @()
    objects = @()
    inventoryCompleteForOrdinaryGroups = $null
    specialContainersInspected = $false
    issues = @()
}
if ($InspectRunning) {
    try {
        if ($PSVersionTable.PSEdition -eq 'Core' -or $PSVersionTable.PSVersion.Major -ne 5) {
            throw 'Use 64-bit Windows PowerShell 5.1 with -STA for GetActiveObject.'
        }
        $app = [Runtime.InteropServices.Marshal]::GetActiveObject($ProgId)
        $report.connection = 'connected'
        $report.applicationVersion = Read-Property $app 'Version' 'application'
        $wanted = $null
        if ($DocumentFullName) { $wanted = [IO.Path]::GetFullPath($DocumentFullName) }
        $target = $null
        $matches = 0
        for ($documentIndex=1; $documentIndex -le $app.Documents.Count; $documentIndex++) {
            $document = $app.Documents.Item($documentIndex)
            $fullName = Read-Property $document 'FullFileName' ('document/' + $documentIndex)
            $documentRows.Add([pscustomobject]@{
                index=$documentIndex; name=(Read-Property $document 'Name' ('document/' + $documentIndex)); fullFileName=$fullName
            })
            if ($wanted -and $fullName -and [IO.Path]::IsPathRooted([string]$fullName)) {
                if ([string]::Equals([IO.Path]::GetFullPath([string]$fullName),$wanted,[StringComparison]::OrdinalIgnoreCase)) {
                    $target = $document
                    $matches++
                }
            }
        }
        if ($wanted -and $matches -ne 1) { throw 'Expected exactly one open document matching the complete path in the connected instance.' }
        if ($target) {
            $report.target = [ordered]@{fullFileName=$wanted}
            foreach ($property in @('Unit','WorldScale','DrawingOriginX','DrawingOriginY')) {
                $report.target[$property] = Read-Property $target $property 'target'
            }
            for ($pageIndex=1; $pageIndex -le $target.Pages.Count; $pageIndex++) {
                $page = $target.Pages.Item($pageIndex)
                $pageRows.Add([pscustomobject]@{
                    index=$pageIndex
                    sizeWidth=(Read-Property $page 'SizeWidth' ('page/' + $pageIndex))
                    sizeHeight=(Read-Property $page 'SizeHeight' ('page/' + $pageIndex))
                })
                if ($Inventory) { Read-Shapes $page.Shapes $pageIndex ('page/' + $pageIndex) 0 }
            }
            if ($Inventory) { $report.inventoryCompleteForOrdinaryGroups = $ordinaryGroupsComplete }
        }
    } catch {
        if ($report.connection -eq 'not_requested') { $report.connection = 'failed' }
        $issues.Add([pscustomobject]@{context='connection_or_document_inspection'; error=$_.Exception.Message})
        $exitStatus = 2
    }
}
$report.documents = @($documentRows.ToArray())
$report.pages = @($pageRows.ToArray())
$report.objects = @($objects.ToArray())
$report.issues = @($issues.ToArray())
if ($issues.Count -gt 0) { $exitStatus = 2 }
$report['status'] = if ($exitStatus -eq 0) { 'ok' } else { 'incomplete_or_failed' }
$json = $report | ConvertTo-Json -Depth 80
if ($OutputPath) {
    # CreateNew is atomic: never overwrite an existing file, even after a race.
    $stream = [IO.File]::Open($OutputPath,[IO.FileMode]::CreateNew,[IO.FileAccess]::Write,[IO.FileShare]::None)
    $writer = New-Object IO.StreamWriter($stream,(New-Object Text.UTF8Encoding($true)))
    try { $writer.Write($json) } finally { $writer.Dispose() }
}
Write-Output $json
exit $exitStatus

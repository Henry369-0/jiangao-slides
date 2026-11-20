#requires -Version 5.1
<#
.SYNOPSIS
Edit one native PowerPoint table cell and save the verified result as a new PPTX.
.DESCRIPTION
Requires desktop Microsoft PowerPoint on Windows. Slide, row, column and
TableIndex are one-based. TableIndex counts top-level native tables in slide
shape order; TableName is an exact, unique top-level native table shape name.
OldText must match the full cell content. Use an empty string for an empty cell.
The source is read from disk. Save any pending edits in PowerPoint first.
.PARAMETER Show
Display the temporary working copy at the target slide during the edit.
The window closes automatically before verification. No manual input is needed.
#>
[CmdletBinding(DefaultParameterSetName='ByName')]
param(
    [Parameter(Mandatory)][string]$InputPath,
    [Parameter(Mandatory)][string]$OutputPath,
    [Parameter(Mandatory)][ValidateRange(1,2147483647)][int]$Slide,
    [Parameter(Mandatory,ParameterSetName='ByName')][ValidateNotNullOrEmpty()][string]$TableName,
    [Parameter(Mandatory,ParameterSetName='ByIndex')][ValidateRange(1,2147483647)][int]$TableIndex,
    [Parameter(Mandatory)][ValidateRange(1,2147483647)][int]$Row,
    [Parameter(Mandatory)][ValidateRange(1,2147483647)][int]$Column,
    [Parameter(Mandatory)][AllowEmptyString()][string]$OldText,
    [Parameter(Mandatory)][AllowEmptyString()][string]$NewText,
    [switch]$Show
)
$ErrorActionPreference = 'Stop'
if ([Environment]::OSVersion.Platform -ne [PlatformID]::Win32NT) {
    throw 'Desktop PowerPoint on Windows is required.'
}
$source = (Get-Item -LiteralPath $InputPath).FullName
$destination = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath)
if (-not [IO.File]::Exists($source) -or [IO.Path]::GetExtension($source) -ine '.pptx' -or
    [IO.Path]::GetExtension($destination) -ine '.pptx') {
    throw 'Input and output must be .pptx files.'
}
if (Test-Path -LiteralPath $destination) { throw 'The output already exists. Choose a new output file.' }
$directory = [IO.Path]::GetDirectoryName($destination)
if (-not [IO.Directory]::Exists($directory)) { throw 'Create the output directory first.' }
$expected = $OldText.Replace("`r`n","`r").Replace("`n","`r")
$replacement = $NewText.Replace("`r`n","`r").Replace("`n","`r")
$byName = $PSCmdlet.ParameterSetName -eq 'ByName'

function Get-TargetCell($deck) {
    if ($Slide -gt $deck.Slides.Count) { throw 'The requested slide is outside this presentation.' }
    $tables = @($deck.Slides.Item($Slide).Shapes | Where-Object { $_.HasTable -eq -1 })
    if ($byName) {
        $matches = @($tables | Where-Object { [string]::Equals($_.Name,$TableName,[StringComparison]::Ordinal) })
        if ($matches.Count -ne 1) { throw 'TableName must identify exactly one native table. Use TableIndex if names repeat.' }
        $table = $matches[0].Table
    } else {
        if ($TableIndex -gt $tables.Count) { throw 'The requested native table index does not exist on this slide.' }
        $table = $tables[$TableIndex-1].Table
    }
    if ($Row -gt $table.Rows.Count -or $Column -gt $table.Columns.Count) {
        throw 'The requested cell is outside this table.'
    }
    return $table.Cell($Row,$Column).Shape.TextFrame.TextRange
}

function Set-CellValue($range, [string]$before, [string]$after) {
    if ([string]::Equals($before,$after,[StringComparison]::Ordinal)) { return }
    # Keep formatting on unchanged runs, e.g. a superscript after an edited value.
    # If a host counts Unicode characters differently, use full-cell replacement.
    if ($before.Length -eq 0 -or $range.Length -ne $before.Length -or
        $before -match '[\uD800-\uDFFF]' -or $after -match '[\uD800-\uDFFF]') {
        $range.Text = $after
        return
    }
    $limit = [Math]::Min($before.Length,$after.Length)
    $prefix = 0
    while ($prefix -lt $limit -and $before[$prefix] -ceq $after[$prefix]) { $prefix++ }
    $suffix = 0
    while ($prefix+$suffix -lt $limit -and $before[$before.Length-1-$suffix] -ceq $after[$after.Length-1-$suffix]) { $suffix++ }
    $oldCount = $before.Length-$prefix-$suffix
    $insert = $after.Substring($prefix,$after.Length-$prefix-$suffix)
    if ($oldCount -gt 0) {
        $range.Characters($prefix+1,$oldCount).Text = $insert
    } elseif ($prefix -eq 0) {
        $null = $range.InsertBefore($insert)
    } else {
        $null = $range.Characters($prefix,1).InsertAfter($insert)
    }
}

$working = Join-Path $directory ('.cell-edit-' + [Guid]::NewGuid().ToString('N') + '.pptx')
$sourceHash = (Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash
$ownsWorking = $false
$deck = $null
try {
    [IO.File]::Copy($source,$working,$false)
    $ownsWorking = $true
    if ((Get-FileHash -LiteralPath $working -Algorithm SHA256).Hash -ne $sourceHash) {
        throw 'The source changed during copying. Retry with the saved source.'
    }
    $workingFile = Get-Item -LiteralPath $working -Force
    $workingFile.IsReadOnly = $false
    $app = New-Object -ComObject PowerPoint.Application
    $withWindow = if ($Show) { -1 } else { 0 }
    $deck = $app.Presentations.Open($working,0,0,$withWindow)
    $target = Get-TargetCell $deck
    if (-not [string]::Equals($target.Text,$expected,[StringComparison]::Ordinal)) {
        throw 'OldText does not match the selected cell. Check the slide, table, row, column and text.'
    }
    if ($Show) {
        $window = $deck.Windows.Item(1)
        $window.Activate()
        $window.View.GotoSlide($Slide)
    }
    Set-CellValue $target $expected $replacement
    $deck.Save()
    $deck.Close()
    $deck = $null
    $deck = $app.Presentations.Open($working,-1,0,0)
    $verified = Get-TargetCell $deck
    if (-not [string]::Equals($verified.Text,$replacement,[StringComparison]::Ordinal)) {
        throw 'The saved cell did not pass verification after reopening.'
    }
    $deck.Close()
    $deck = $null
    if ((Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash -ne $sourceHash) {
        throw 'The source changed during the edit. Retry with the latest saved file.'
    }
    # Same-directory move publishes only a completed file and refuses overwrites.
    [IO.File]::Move($working,$destination)
    $ownsWorking = $false
    [pscustomobject]@{OutputFile=[IO.Path]::GetFileName($destination); Slide=$Slide; Row=$Row; Column=$Column; Verified=$true}
} finally {
    if ($null -ne $deck) { $deck.Saved=-1; $deck.Close() }
    if ($ownsWorking -and [IO.File]::Exists($working)) { [IO.File]::Delete($working) }
}

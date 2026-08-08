# ============================================================
# PowerShellTools Project Launcher
#
# Usage:
#   .\PSTP.ps1 <Tool> [tool parameters...]
#
# Examples:
#   .\PSTP.ps1 Create
#   .\PSTP.ps1 Release -Zip
#   .\PSTP.ps1 Release -NoBump
#   .\PSTP.ps1 Language
#
# Tool naming convention:
#   <Tool>
#     -> PowerShellTools\Project\<Tool>\Project<Tool>.ps1
#
# Example:
#   Release
#     -> PowerShellTools\Project\Release\ProjectRelease.ps1
# ============================================================

if ($args.Count -eq 0) {
    Write-Host ""
    Write-Host "[ERROR] No project tool specified." -ForegroundColor Red
    Write-Host ""
    Write-Host "Usage:"
    Write-Host "  .\PSTP.ps1 <Tool> [tool parameters...]"
    Write-Host ""
    exit 1
}

$toolName = [string]$args[0]

if ([string]::IsNullOrWhiteSpace($toolName)) {
    Write-Host ""
    Write-Host "[ERROR] Project tool name cannot be empty." -ForegroundColor Red
    Write-Host ""
    exit 1
}

# Everything after the tool name belongs to the tool itself.
$toolArgs = @()

if ($args.Count -gt 1) {
    $toolArgs = @($args[1..($args.Count - 1)])
}

# The folder containing PSTP.ps1 is the project folder.
$projectFolder = $PSScriptRoot

# Search upwards for the PowerShellTools repository.
$searchFolder = $projectFolder
$toolScript = $null

while ($searchFolder) {

    $candidate = Join-Path $searchFolder `
        "PowerShellTools\Project\$toolName\Project$toolName.ps1"

    if (Test-Path -LiteralPath $candidate -PathType Leaf) {
        $toolScript = $candidate
        break
    }

    $parent = Split-Path $searchFolder -Parent

    if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $searchFolder) {
        break
    }

    $searchFolder = $parent
}

if (-not $toolScript) {
    Write-Host ""
    Write-Host "[ERROR] Project tool '$toolName' was not found." -ForegroundColor Red
    Write-Host ""
    Write-Host "Expected:"
    Write-Host "  PowerShellTools\Project\$toolName\Project$toolName.ps1"
    Write-Host ""
    exit 1
}

# The project tool is responsible for validating its own parameters.
& $toolScript -ProjectFolder $projectFolder @toolArgs

exit $LASTEXITCODE
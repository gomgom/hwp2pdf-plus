# hwp2pdf+ build script
#   powershell -ExecutionPolicy Bypass -File build.ps1
#
# Builds a single-file WPF exe with the C# compiler shipped in .NET Framework
# (no NuGet, no external dependencies).
#
# NOTE: kept ASCII-only on purpose - Windows PowerShell reads BOM-less .ps1 files
# using the ANSI codepage, which would corrupt non-ASCII text.
[CmdletBinding()]
param(
    [string]$Output = "hwp2pdf-plus.exe"
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $MyInvocation.MyCommand.Definition
Push-Location $root
try {
    $fw = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319'
    if (-not (Test-Path $fw)) { $fw = Join-Path $env:WINDIR 'Microsoft.NET\Framework\v4.0.30319' }

    $csc = Join-Path $fw 'csc.exe'
    if (-not (Test-Path $csc)) {
        throw "csc.exe not found at $csc - .NET Framework 4.x is required."
    }

    # Resolve references: framework folder first, then its WPF subfolder.
    $searchDirs = @($fw, (Join-Path $fw 'WPF'))
    $needed = @(
        'System.dll', 'System.Core.dll', 'System.Xml.dll',
        'System.Drawing.dll', 'System.Windows.Forms.dll',
        'System.Xaml.dll', 'WindowsBase.dll', 'PresentationCore.dll', 'PresentationFramework.dll'
    )

    $refs = @()
    foreach ($name in $needed) {
        $found = $null
        foreach ($d in $searchDirs) {
            $p = Join-Path $d $name
            if (Test-Path $p) { $found = $p; break }
        }
        if (-not $found) {
            throw "Reference assembly not found: $name (searched: $($searchDirs -join '; '))"
        }
        $refs += "/reference:$found"
    }

    $cscArgs = @(
        '/nologo'
        '/target:winexe'
        '/codepage:65001'
        '/optimize+'
        '/win32icon:res\app.ico'
        '/resource:res\app.ico,app.ico'
        '/resource:ui\MainWindow.xaml,MainWindow.xaml'
        "/out:$Output"
    ) + $refs + @('Program.cs')

    Write-Host "Building $Output ..."
    & $csc $cscArgs
    if ($LASTEXITCODE -ne 0) { throw "Compilation failed (exit $LASTEXITCODE)" }

    $info = [System.Diagnostics.FileVersionInfo]::GetVersionInfo((Resolve-Path $Output).Path)
    $size = [math]::Round((Get-Item $Output).Length / 1KB, 1)
    Write-Host ("Build OK: {0} ({1} {2}, {3} KB)" -f $Output, $info.ProductName, $info.FileVersion, $size)
}
finally {
    Pop-Location
}

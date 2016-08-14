#requires -Version 3
# Title: Fix-AppvMsiPackage.ps1
# Description: Applies AppvMsiFixer.mst to MSI and cleans up some additional items
# Author: Dan Gough
# Version: 1.0
# Dependencies: WixToolset.Dtf.WindowsInstaller.dll and PowerShell v3.0

param
(
    [Parameter(Mandatory = $true, HelpMessage = 'Path to MSI package or folder')]
    [System.String]
    $Path
)

Add-Type -Path ((Split-Path -Parent -Path $PSCommandPath) + '\WixToolset.Dtf.WindowsInstaller.dll')
function Open-MsiDatabase
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true,
        HelpMessage = 'Path to MSI package')]
        [System.String]
        $DatabasePath,
        [Parameter(Position = 1)]
        [WixToolset.Dtf.WindowsInstaller.DatabaseOpenMode]
        $TransactMode = 'ReadOnly'
    )

    $Database = New-Object -TypeName WixToolset.Dtf.WindowsInstaller.Database -ArgumentList $DatabasePath, $TransactMode
    Write-Output $Database
}

$Packages = @()

If (Test-Path $Path -PathType Leaf)
{
    If ([System.IO.Path]::GetExtension($Path) -eq '.msi')
    {
        $Packages = $Path
    }
    Else
    {
        Write-Error 'Please specift a valid .msi file or directory path.'
    }
}
Elseif (Test-Path $Path -PathType Container)
{
    $Packages = (Get-ChildItem $Path -Filter *.msi -Recurse).Fullname
    If ($Packages.Count -eq 0)
    {
        Write-Error "No MSI packages found in '$Path'."
    }
}
Else
{
    Write-Error "Path '$Path' not found."
}

$Packages | ForEach-Object {
    Write-Output "Processing $_..."
    If ((Test-Path "$_.bak") -eq $false) {
        Copy-Item -Path $_ -Destination "$_.bak"
    }
    $MSI = Open-MsiDatabase -DatabasePath $_ -TransactMode Transact
    $MSI.ApplyTransform((Split-Path -Parent -Path $PSCommandPath) + '\Fix-AppvMsiPackage.mst')
    $MSI.ApplyTransform((Split-Path -Parent -Path $PSCommandPath) + '\Fix-AppvMsiPackage_1607x86.mst')
    $MSI.Execute("DELETE FROM Property WHERE Property = 'MstVersion'")
    $MSI.Commit()
    $MSI.Dispose()
}
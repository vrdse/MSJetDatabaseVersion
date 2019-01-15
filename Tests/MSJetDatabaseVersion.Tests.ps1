$TestPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$RepositoryPath = Split-Path -Path $TestPath
$ModulePath = "$RepositoryPath\MSJetDatabaseVersion"
# $ModuleScript = "$ModulePath\MSJetDatabaseVersion.psm1"
$ModuleManifest = "$ModulePath\MSJetDatabaseVersion.psd1"

# test the module manifest - exports the right functions, processes the right formats, and is generally correct
Describe "Manifest" {

    $ManifestHash = Invoke-Expression (Get-Content $ModuleManifest -Raw)

    It "has a valid manifest" {
        {
            $null = Test-ModuleManifest -Path $ModuleManifest -ErrorAction Stop -WarningAction SilentlyContinue
        } | Should Not Throw
    }

    It "has a valid root module" {
        $ManifestHash.RootModule | Should Be "MSJetDatabaseVersion.psm1"
    }

    It "has a valid Description" {
        $ManifestHash.Description | Should Not BeNullOrEmpty
    }

    It "has a valid guid" {
        $ManifestHash.Guid | Should Match '[0-9A-Fa-f]{8}[-][0-9A-Fa-f]{4}[-][0-9A-Fa-f]{4}[-][0-9A-Fa-f]{4}[-][0-9A-Fa-f]{12}'
    }

    It "has a valid copyright" {
        $ManifestHash.CopyRight | Should Not BeNullOrEmpty
    }

    Context "exports all public functions" {

        # All functions that are defined in the Manifest
        $ExFunctions = $ManifestHash.FunctionsToExport
        # All functions that are supposed to get exported (Files in Public folder)
        $FunctionFiles = Get-ChildItem "$ModulePath\Public" -Filter *.ps1 | Select-Object -ExpandProperty BaseName
        $FunctionNames = $FunctionFiles

        # Test if all public functions are specified in the Manifest
        foreach ($FunctionName in $FunctionNames)
        {
            It "exports $FunctionName" {
                $ExFunctions -contains $FunctionName | Should Be $true
            }
        }
    }
}



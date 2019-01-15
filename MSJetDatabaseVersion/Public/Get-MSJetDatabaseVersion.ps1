<#
.SYNOPSIS
    Identifies the Version of Microsoft Access JET DBs.
.DESCRIPTION
    Identifies the Version of Microsoft Access JET DBs by looking at its file header.
.EXAMPLE
    PS C:\> Get-MSJetDatabaseVersion -Path C:\database.mdb
    Returns the JET version of C:\database.mdb
.EXAMPLE
    PS C:\> Get-ChildItem -Path C:\ -Recurse -Force -ErrorAction SilentlyContinue -Filter *.mdb | Get-MSJetDatabaseVersion
    Search local C drive for mdb-Files and get their JET version.
.NOTES
    Author: VRDSE
#>
function Get-MSJetDatabaseVersion {
    [CmdletBinding()]
    param (
        # Path
        [Parameter(
            Mandatory = $true,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [Alias('FullName')]
        [string]
        $Path
    )
     
    begin {
        # Source: http://jabakobob.net/mdb/first-page.html
        $JetVersionHashTable = @{
            '00000000' = 'Access 97 (Jet 3)'
            '01000000' = 'Access 2000, 2002/2003 (Jet 4)'
        }
    }
     
    process {
         
        Write-Verbose -Message ('Proccessing {0}' -f $Path)
 
        [Byte[]]$FileHeader = Get-Content -Path $Path -TotalCount 24 -Encoding Byte
        $FileFormat = [System.Text.Encoding]::ASCII.GetString($FileHeader[4..18])

        if ($FileFormat -eq 'Standard Jet DB') {

            Write-Verbose -Message 'Standard Jet DB found. Test for JET file version'
            $JetVersionHexAsString = $($FileHeader[20..24] | ForEach-Object { $_.ToString('X2') }) -join ''

            if ($JetVersionHashTable[$JetVersionHexAsString]) {

                [PSCustomObject]@{
                    'Path'       = $Path
                    'JetVersion' = $JetVersionHashTable[$JetVersionHexAsString]
                }

            }
            else {

                Write-Verbose -Message ('Standard Jet DB found, but unknown format (0x{0})' -f $JetVersionHexAsString)

            }
            
        }
        else {
            
            Write-Verbose -Message 'No Standard Jet DB'

        }

    }
     
    end {
    }
}
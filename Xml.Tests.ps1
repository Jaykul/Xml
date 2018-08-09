$scriptroot = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
Import-Module (Join-Path $scriptroot 'XML\XML.psm1') -Force

$PSVersion = $PSVersionTable.PSVersion.Major
Describe 'New-XDocument' {
    Context "Strict mode PS$PSVersion" {
        Set-StrictMode -Version latest
        It "Creates a simple document with no namespaces" {
            New-XDocument -root 'movies' -Content {
                movie { title { 'Aliens' } }
            }
        }
    }
}
$exportPath = "C:\GPOReports_JSON"
New-Item -Path $exportPath -ItemType Directory -Force | Out-Null

$allGPOs = Get-GPO -All
$fullJson = @()

foreach ($gpo in $allGPOs) {
    Write-Host "Processing GPO: $($gpo.DisplayName)"

    # Get WMI Filter if it exists
    $wmiFilter = Get-GPOWmiFilter -Guid $gpo.Id -ErrorAction SilentlyContinue

    # Get linked targets (OU, domain, site)
    $gpoLinks = Get-GPOLinkedTargets -GPO $gpo.DisplayName

    # Export GPO report to memory
    $reportXml = Get-GPOReport -Guid $gpo.Id -ReportType Xml
    [xml]$gpoXml = $reportXml

    function Extract-Settings {
        param (
            [xml]$gpoNode,
            [string]$scope # "Computer" or "User"
        )

        $results = @()

        foreach ($extension in $gpoNode.$scope.ExtensionData.Extension) {
            switch ($extension.Name) {

                "RegistrySettings" {
                    foreach ($reg in $extension.RegistrySettings.Registry) {
                        $results += [PSCustomObject]@{
                            GPOName     = $gpo.DisplayName
                            AppliesTo   = $scope
                            Category    = "Registry"
                            Path        = $reg.Key
                            Name        = $reg.ValueName
                            Type        = $reg.Type
                            Data        = $reg.Value
                            Hive        = $reg.Hive
                            WmiFilter   = $wmiFilter?.Query
                            LinkedTo    = $gpoLinks -join "; "
                        }
                    }
                }

                "Scripts" {
                    foreach ($script in $extension.Scripts.Script) {
                        $results += [PSCustomObject]@{
                            GPOName     = $gpo.DisplayName
                            AppliesTo   = $scope
                            Category    = "Scripts"
                            Script      = $script.Script
                            Parameters  = $script.Parameters
                            WmiFilter   = $wmiFilter?.Query
                            LinkedTo    = $gpoLinks -join "; "
                        }
                    }
                }

                "SecuritySettings" {
                    $secSettings = $extension.SecuritySettings
                    if ($secSettings) {
                        $results += [PSCustomObject]@{
                            GPOName   = $gpo.DisplayName
                            AppliesTo = $scope
                            Category  = "Security"
                            Setting   = "Various (password, lockout, audit, etc.)"
                            Details   = $secSettings.InnerXml
                            WmiFilter = $wmiFilter?.Query
                            LinkedTo  = $gpoLinks -join "; "
                        }
                    }
                }

                default {
                    $results += [PSCustomObject]@{
                        GPOName   = $gpo.DisplayName
                        AppliesTo = $scope
                        Category  = $extension.Name
                        Details   = $extension.InnerXml
                        WmiFilter = $wmiFilter?.Query
                        LinkedTo  = $gpoLinks -join "; "
                    }
                }
            }
        }

        return $results
    }

    $computerSettings = Extract-Settings -gpoNode $gpoXml.GPO -scope "Computer"
    $userSettings     = Extract-Settings -gpoNode $gpoXml.GPO -scope "User"

    $fullJson += $computerSettings + $userSettings
}

# Export final JSON
$jsonFile = "$exportPath\All_GPOs_Enhanced_WithWMI.json"
$fullJson | ConvertTo-Json -Depth 5 | Out-File -Encoding UTF8 -FilePath $jsonFile

Write-Host "`nExport completed. File saved to: $jsonFile"

# Helper function to find linked targets
function Get-GPOLinkedTargets {
    param ([string]$GPO)

    $targets = @()
    $ouLinks = Get-ADOrganizationalUnit -Filter * | ForEach-Object {
        $inherit = Get-GPInheritance -Target $_.DistinguishedName
        $link = $inherit.GpoLinks | Where-Object { $_.DisplayName -eq $GPO }
        if ($link) { $_.DistinguishedName }
    }
    $domainLink = Get-GPInheritance -Target "DC=$(Get-ADDomain).DNSRoot" |
        Select-Object -ExpandProperty GpoLinks |
        Where-Object { $_.DisplayName -eq $GPO } |
        ForEach-Object { "Domain: $($_.DisplayName)" }

    $targets += $ouLinks
    $targets += $domainLink
    return $targets
}

# Directory to store output
$exportPath = "C:\GPOReports_JSON"
New-Item -Path $exportPath -ItemType Directory -Force | Out-Null

$allGPOs = Get-GPO -All
$fullJson = @()

foreach ($gpo in $allGPOs) {
    Write-Host "Processing GPO: $($gpo.DisplayName)"

    # Export XML report to memory
    $reportXml = Get-GPOReport -Guid $gpo.Id -ReportType Xml
    [xml]$gpoXml = $reportXml

    # Function to extract settings (Computer/User) from RegistrySettings, Scripts, Security, etc.
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
                        }
                    }
                }

                default {
                    $results += [PSCustomObject]@{
                        GPOName   = $gpo.DisplayName
                        AppliesTo = $scope
                        Category  = $extension.Name
                        Details   = $extension.InnerXml
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

# Export the final JSON
$jsonFile = "$exportPath\All_GPOs_Enhanced.json"
$fullJson | ConvertTo-Json -Depth 5 | Out-File -Encoding UTF8 -FilePath $jsonFile

Write-Host "`nExport completed. JSON file located at: $jsonFile"

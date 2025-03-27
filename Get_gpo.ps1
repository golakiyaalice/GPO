# Create export folder
$exportPath = "C:\GPOReports_JSON"
New-Item -Path $exportPath -ItemType Directory -Force

$allGPOs = Get-GPO -All
$jsonOutput = @()

foreach ($gpo in $allGPOs) {
    $reportXmlPath = "$exportPath\$($gpo.DisplayName).xml"
    $report = Get-GPOReport -Guid $gpo.Id -ReportType Xml

    # Save raw XML report
    $report | Out-File -FilePath $reportXmlPath

    # Load XML for parsing
    [xml]$xml = $report

    # Extract registry-based settings (computer + user)
    $settings = @()
    foreach ($setting in $xml.GPO.Computer.ExtensionData.Extension | Where-Object { $_.Name -eq "RegistrySettings" }) {
        $setting.RegistrySettings.Registry | ForEach-Object {
            $settings += [PSCustomObject]@{
                Path = $_.Key
                Name = $_.ValueName
                Type = $_.Type
                Data = $_.Value
                Hive = $_.Hive
                AppliesTo = "Computer"
            }
        }
    }
    foreach ($setting in $xml.GPO.User.ExtensionData.Extension | Where-Object { $_.Name -eq "RegistrySettings" }) {
        $setting.RegistrySettings.Registry | ForEach-Object {
            $settings += [PSCustomObject]@{
                Path = $_.Key
                Name = $_.ValueName
                Type = $_.Type
                Data = $_.Value
                Hive = $_.Hive
                AppliesTo = "User"
            }
        }
    }

    # Prepare JSON-friendly object
    $jsonOutput += [PSCustomObject]@{
        GPOName = $gpo.DisplayName
        GPOID = $gpo.Id
        CreatedTime = $gpo.CreationTime
        ModifiedTime = $gpo.ModificationTime
        Settings = $settings
        Links = (Get-GPOReport -Guid $gpo.Id -ReportType Html | Select-String -Pattern "Linked to:.*?" -AllMatches).Matches.Value
        OSApplicability = "All Windows OS (GPOs are generic unless WMI filters are applied)"
    }
}

# Save full JSON for Ansible
$jsonFile = "$exportPath\All_GPOs_Ansible_Format.json"
$jsonOutput | ConvertTo-Json -Depth 5 | Out-File -FilePath $jsonFile -Encoding UTF8
Write-Host "Exported JSON to $jsonFile"

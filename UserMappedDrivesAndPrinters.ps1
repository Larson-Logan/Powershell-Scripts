# Get the current user's desktop path
$desktopPath = [System.IO.Path]::Combine([System.Environment]::GetFolderPath("Desktop"), "UserMappedDrivesAndPrinters.csv")

# Initialize an empty array to hold the results
$results = @()
$mappedPrinterNames = @()

# Get the list of user profiles
$userProfiles = Get-WmiObject -Class Win32_UserProfile | Where-Object { $_.Special -eq $false }

foreach ($profile in $userProfiles) {
    try {
        $profilePath = $profile.LocalPath
        $sid = $profile.SID

        # Get the user name, checking for Azure AD users as well
        $userName = (Get-WmiObject -Class Win32_UserAccount | Where-Object { $_.SID -eq $sid }).Name
        if (-not $userName) {
            $userName = (Get-WmiObject -Class Win32_UserAccount -Namespace "Root\Microsoft\Identity\Providers" | Where-Object { $_.SID -eq $sid }).Name
        }

        $lastUseTime = $profile.LastUseTime

        Write-Host "`nUser: $userName"

        # Initialize flags to check if mapped drives and printers are found
        $hasMappedDrives = $false
        $hasMappedPrinters = $false
        $hasLocalPrinters = $false

        # Check for mapped drives
        $mappedDrivesKey = "Registry::HKEY_USERS\$($profile.SID)\Network"
        if (Test-Path $mappedDrivesKey) {
            $mappedDrives = Get-ChildItem -Path $mappedDrivesKey
            if ($mappedDrives) {
                Write-Host "  Mapped Drives:"
                $hasMappedDrives = $true
                foreach ($drive in $mappedDrives) {
                    $driveLetter = $drive.PSChildName
                    $remotePath = (Get-ItemProperty -Path "$mappedDrivesKey\$driveLetter").RemotePath
                    Write-Host "    ${driveLetter}"
                    $results += [pscustomobject]@{
                        UserName    = $userName
                        ProfilePath = $profilePath
                        LastUseTime = $lastUseTime
                        Type        = "Mapped Drive"
                        Identifier  = $driveLetter
                        Details     = $remotePath
                    }
                }
            }
        }

        if (-not $hasMappedDrives) {
            Write-Host "  No Mapped Drives"
            $results += [pscustomobject]@{
                UserName    = $userName
                ProfilePath = $profilePath
                LastUseTime = $lastUseTime
                Type        = "Mapped Drive"
                Identifier  = "None"
                Details     = "No Mapped Drives"
            }
        }

        # Check for mapped printers in the registry
        $printersKey = "Registry::HKEY_USERS\$($profile.SID)\Software\Microsoft\Windows NT\CurrentVersion\PrinterPorts"
        if (Test-Path $printersKey) {
            $mappedPrinters = Get-ItemProperty -Path $printersKey
            if ($mappedPrinters) {
                Write-Host "  Mapped Printers:"
                $hasMappedPrinters = $true
                foreach ($printer in $mappedPrinters.PSObject.Properties) {
                    if ($printer.Name -notin @('PSPath', 'PSParentPath', 'PSChildName', 'PSProvider')) {
                        $printerDetails = $printer.Value -split ","
                        $port = $printerDetails[0]
                        Write-Host "    $($printer.Name) (Port: $port)"
                        $results += [pscustomobject]@{
                            UserName    = $userName
                            ProfilePath = $profilePath
                            LastUseTime = $lastUseTime
                            Type        = "Mapped Printer"
                            Identifier  = $printer.Name
                            Details     = "Port: $port"
                        }
                        $mappedPrinterNames += $printer.Name
                    }
                }
            }
        }

        if (-not $hasMappedPrinters) {
            Write-Host "  No Mapped Printers"
            $results += [pscustomobject]@{
                UserName    = $userName
                ProfilePath = $profilePath
                LastUseTime = $lastUseTime
                Type        = "Mapped Printer"
                Identifier  = "None"
                Details     = "No Mapped Printers"
            }
        }

        # Check for local printers and their ports
        $printers = Get-WmiObject -Query "SELECT * FROM Win32_Printer"
        if ($printers) {
            Write-Host "  Local Printers:"
            $hasLocalPrinters = $true
            foreach ($printer in $printers) {
                if ($mappedPrinterNames -notcontains $printer.Name) {
                    $portName = $printer.PortName
                    $printerType = if ($portName -match "^(IP|TCP|LPT|USB)_") { "Network Printer" } else { "Local Printer" }
                    Write-Host "    $($printer.Name) (Port: ${portName})"
                    $results += [pscustomobject]@{
                        UserName    = $userName
                        ProfilePath = $profilePath
                        LastUseTime = $lastUseTime
                        Type        = $printerType
                        Identifier  = $printer.Name
                        Details     = "Port: $portName"
                    }
                }
            }
        }

        if (-not $hasLocalPrinters) {
            Write-Host "  No Local Printers"
            $results += [pscustomobject]@{
                UserName    = $userName
                ProfilePath = $profilePath
                LastUseTime = $lastUseTime
                Type        = "Printer"
                Identifier  = "None"
                Details     = "No Local Printers"
            }
        }

    } catch {
        Write-Host "Error processing user: $userName"
        Write-Host $_.Exception.Message
        $results += [pscustomobject]@{
            UserName    = $userName
            ProfilePath = $profilePath
            LastUseTime = $lastUseTime
            Type        = "Error"
            Identifier  = "N/A"
            Details     = $_.Exception.Message
        }
    }
}

# Export results to CSV
$results | Export-Csv -Path $desktopPath -NoTypeInformation -Force

Write-Host "Results have been exported to $desktopPath"
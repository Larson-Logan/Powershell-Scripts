# Get the current user's desktop path
$desktopPath = [System.IO.Path]::Combine([System.Environment]::GetFolderPath("Desktop"), "UserMappedDrivesAndPrinters.csv")

# Initialize an empty array to hold the results
$results = @()
$mappedPrinterNames = @()

# Define an array of additional port prefixes
# $additionalPorts = @()

# Combine the additional ports with the existing ones
# $allPorts = "^(IP|TCP|WSD-|IPP|HTTP|HTTPS|FTP|SSH|SMTP|DNS|LPR|RAW|SMB)|^\d{1,3}(\.\d{1,3}){3}$" + ($additionalPorts -join "|")

function Get-NetworkInfo {
    param (
        [string]$ipAddress
    )

    $url = "https://whois.arin.net/rest/ip/$ipAddress"
    try {
        $response = Invoke-RestMethod -Uri $url -Method Get -Headers @{Accept = "application/json"}
        return $response
    } catch {
        Write-Output "Error querying ARIN for IP: $ipAddress"
        return $null
    }
}

function Test-NetworkPort {
    param (
        [string]$portName
    )

    if ($portName -match "^(IP|TCP|WSD-|IPP_|HTTP|HTTPS|FTP|NET|SSH|SMTP|DNS|LPR|RAW|SMB)") {
        return $true
    } elseif ($portName -match "^\d{1,3}(\.\d{1,3}){3}$") {
        $networkInfo = Get-NetworkInfo -ipAddress $portName
        if ($networkInfo) {
            return $true
        }
    }
    return $false
}

# Get the list of user profiles
$userProfiles = Get-WmiObject -Class Win32_UserProfile | Where-Object { $_.Special -eq $false }

foreach ($profile in $userProfiles) {
    try {
        $profilePath = $profile.LocalPath
        $sid = $profile.SID

        # Get the user name, checking for Azure AD users as well
        $userName = (Get-WmiObject -Class Win32_UserAccount | Where-Object { $_.SID -eq $sid }).Name
        if (-not $userName) {
            if (Get-WmiObject -Namespace "Root\Microsoft\Identity\Providers" -Class __Namespace -ErrorAction SilentlyContinue) {
                $userName = (Get-WmiObject -Class Win32_UserAccount -Namespace "Root\Microsoft\Identity\Providers" | Where-Object { $_.SID -eq $sid }).Name
            } else {
                Write-Host "Namespace 'Root\Microsoft\Identity\Providers' not found. Using profile path to determine user name."
                $userName = $profilePath.Split('\')[-1]
            }
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

        # Check for mapped printers using WScript.Network
        $network = New-Object -ComObject WScript.Network
        $printerConnections = $network.EnumPrinterConnections()
        $printerCount = $printerConnections.Count()

        if ($printerCount -gt 0) {
            Write-Host "  Mapped Printers:"
            $hasMappedPrinters = $true
            for ($i = 0; $i -lt $printerCount; $i += 2) {
                $portName = $printerConnections.Item($i)
                $printerName = $printerConnections.Item($i + 1)
                $printer = (Get-Printer)[$i]
                if ($printer) {
                    $printerType = if (Test-NetworkPort -portName $portName) { "Network Printer" } else { "Local Printer" }
                    Write-Host "    $printerName (Port: $portName)"
                    $results += [pscustomobject]@{
                        UserName    = $userName
                        ProfilePath = $profilePath
                        LastUseTime = $lastUseTime
                        Type        = $printerType
                        Identifier  = $printerName
                        Details     = "Port: $portName"
                    }
                    $mappedPrinterNames += $printerName
                } else {
                    Write-Host "    $printerName (Port: $portName)"
                }
            }
        }

        # Check for mapped printers in the registry
        $printersKey = "Registry::HKEY_USERS\$($profile.SID)\Software\Microsoft\Windows NT\CurrentVersion\PrinterPorts"
        if (Test-Path $printersKey) {
            $mappedPrinters = Get-ItemProperty -Path $printersKey
            if ($mappedPrinters) {
                Write-Host "  Mapped Printers (Registry):"
                foreach ($printer in $mappedPrinters.PSObject.Properties) {
                    if ($printer.Name -notin @('PSPath', 'PSParentPath', 'PSChildName', 'PSProvider') -and $mappedPrinterNames -notcontains $printer.Name) {
                        $printerObject = Get-Printer | Where-Object { $_.Name -eq $printer.Name }
                        if ($printerObject) {
                            $portName = $printerObject.PortName
                            $printerType = if (Test-NetworkPort -portName $portName) { "Network Printer" } else { "Local Printer" }
                            Write-Host "    $($printer.Name) (Port: $portName)"
                            $results += [pscustomobject]@{
                                UserName    = $userName
                                ProfilePath = $profilePath
                                LastUseTime = $lastUseTime
                                Type        = $printerType
                                Identifier  = $printer.Name
                                Details     = "Port: $portName"
                            }
                            $mappedPrinterNames += $printer.Name
                        } else {
                            Write-Host "Printer $($printer.Name) not found using Get-Printer."
                        }
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
                    $portName = $printer.PortName # Retrieve the port name from the printer object
                    $printerType = if (Test-NetworkPort -portName $portName) { "Network Printer" } else { "Local Printer" }
                    Write-Host "    $($printer.Name) (Port: $portName)"
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
# Get the SID of the user
$user = "logan"
$sid = $profile.SID
# Query the registry for printers assigned to the user
$printers = Get-ChildItem "Registry::HKEY_USERS\$sid\Printers\Connections"

# Display the printers
$printers | ForEach-Object {
    Write-Output $_.Name "test"
}

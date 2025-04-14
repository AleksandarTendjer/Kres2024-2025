# Query WMI for COM ports
$comPorts = Get-WmiObject -Query "SELECT * FROM Win32_PnPEntity WHERE DeviceID LIKE '%COM%'" | 
            Where-Object {$_.Name -match 'COM'}

# Display the COM ports found
if ($comPorts) {
    foreach ($port in $comPorts) {
        $portName = $port.Name
        $portDeviceID = $port.DeviceID
        Write-Output "COM Port Name: $portName"
        Write-Output "COM Port DeviceID: $portDeviceID"
        Write-Output ""
    }
} else {
    Write-Output "No COM ports found."
}


#Arduino alive
# Define the serial port settings
$portName = "COM11"  # Replace with your Arduino's COM port
$baudRate = 9600    # Match with the baud rate set in your Arduino sketch

# Create a .NET serial port object
$serialPort = New-Object System.IO.Ports.SerialPort
$serialPort.PortName = $portName
$serialPort.BaudRate = $baudRate

# This will give you the directory of the script
$scriptDirectory = $PSScriptRoot

# Define the file name
$fileWelcome = "welcome.html"

# Define the file name for the kiosk website
$fileLocal = "video1.html"
$filePage = Join-Path -Path $scriptDirectory -ChildPath $fileLocal

# Define the file name
$fileSaver = "VideoScreensaver-1.0\VideoScreensaver.scr"

# Combine the directory and file name to get the full path
$fullWelcome = Join-Path -Path $scriptDirectory -ChildPath $fileWelcome

# Combine the directory and file name to get the full path
$fullSaver = Join-Path -Path $scriptDirectory -ChildPath $fileSaver


# Function to preload Chrome with two tabs
function Preload-Chrome {
    $chromeProcess = Get-Process -Name "chrome" -ErrorAction SilentlyContinue
    if (-not $chromeProcess) {
        Write-Host "Preloading Chrome with welcome.html in fullscreen..."
        $chromePath = "C:\Program Files\Google\Chrome\Application\chrome.exe"
        Start-Process $chromePath "--kiosk file:///$fullWelcome --new-window --start-fullscreen"
        Start-Sleep -Seconds 5
        Write-Host "Opening kiosk website in Chrome..."
        Start-Process $chromePath "--kiosk $filePage" -NoNewWindow 
        Start-Sleep -Seconds 1
    } else {
        Write-Host "Chrome is already running."
    }
}

# Function to show the intro screen by switching to the first tab in Chrome
function Show-GraphicalIntro {
    $serialPort.Open()
    $serialPort.WriteLine("1")
    $serialPort.Close()
    Write-Host "Showing graphical intro screen..."
    $chromeProcess = Get-Process -Name "chrome" -ErrorAction SilentlyContinue
    if ($chromeProcess) {
        $wshell = New-Object -ComObject wscript.shell
        $wshell.AppActivate($chromeProcess[0].MainWindowTitle)
        Start-Sleep -Milliseconds 1000
        $wshell.SendKeys("^1") # Switch to the first tab
    }
}
# Function to show the kiosk website by switching to the second tab in Chrome
function Show-KioskWebsite {
    Write-Host "Showing kiosk website..."
    $chromeProcess = Get-Process -Name "chrome" -ErrorAction SilentlyContinue
    if ($chromeProcess) {
        $wshell = New-Object -ComObject wscript.shell
        $wshell.AppActivate($chromeProcess[0].MainWindowTitle)
        Start-Sleep -Milliseconds 500
        $wshell.SendKeys("^2")  # Switch to the second tab
    }
}


# Function to mute the system volume
function Mute-Volume {
    Write-Host "Muting system volume..."
    nircmd.exe mutesysvolume 1
}

# Function to unmute the system volume
function Unmute-Volume {
    Write-Host "Unmuting system volume..."
    nircmd.exe mutesysvolume 0
}

# Function to trigger the screensaver
function Trigger-Screensaver {
    # Switch to the first tab in Chrome before activating the screensaver
    $chromeProcess = Get-Process -Name "chrome" -ErrorAction SilentlyContinue
    if ($chromeProcess) {
        $wshell = New-Object -ComObject wscript.shell
        $wshell.AppActivate($chromeProcess[0].MainWindowTitle)
        Start-Sleep -Milliseconds 500
        $wshell.SendKeys("^1") # Switch to the first tab
    }
    Write-Host "Triggering screensaver..."
    $screensaverPath = "$fullSaver"
    if (Test-Path $screensaverPath) {
        # Mute system volume
        Mute-Volume
        Start-Process $screensaverPath
    } else {
        Write-Host "Screensaver path not found: $screensaverPath"
    }
}

# Function to deactivate the screensaver
function Deactivate-Screensaver {
   Write-Host "Deactivating screensaver..."
    
    # Close the screensaver process if it's running
    $screensaverProcess = Get-Process -Name "VideoScreensaver" -ErrorAction SilentlyContinue
    if ($screensaverProcess) {
        Stop-Process -Name "VideoScreensaver" -Force
    }

    # Unmute system volume
    Unmute-Volume
     # Wait a moment before showing the graphical intro
    Start-Sleep -Milliseconds 500
    Show-GraphicalIntro
    Start-Sleep -Milliseconds 3000
    Show-KioskWebsite
}

# Function to count connected USB devices
function Count-USBDevices {
    Write-Host "Counting USB devices..."
    Get-WmiObject -Query "SELECT * FROM Win32_USBControllerDevice" | Measure-Object | Select-Object -ExpandProperty Count
}

# Preload Chrome with two tabs
Preload-Chrome

# Main loop
Write-Host "Entering main loop..."
$previousDeviceCount = Count-USBDevices
# Main loop
Write-Host "Entering main loop..."
$previousDeviceCount = Count-USBDevices

while ($true) {
    Write-Host "Inside loop, waiting 500 ms..."
    Start-Sleep -Milliseconds 100
    $currentDeviceCount = Count-USBDevices

    if ($currentDeviceCount -lt $previousDeviceCount) {
        Write-Host "USB device count decreased. Triggering screensaver and sending '0' to serial port..."
        Trigger-Screensaver
        $serialPort.Open()
        $serialPort.WriteLine("0")
        $serialPort.Close()
    } elseif ($currentDeviceCount -gt $previousDeviceCount) {
        Write-Host "USB device count increased. Deactivating screensaver and sending '1' to serial port..."
        Deactivate-Screensaver
    }

    $previousDeviceCount = $currentDeviceCount
}


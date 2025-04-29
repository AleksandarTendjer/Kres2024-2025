# Load the AudioDeviceCmdlets module
Import-Module AudioDeviceCmdlets

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
$fileSaver = "videosaver.html"

# Combine the directory and file name to get the full path
$fullWelcome = Join-Path -Path $scriptDirectory -ChildPath $fileWelcome

# Combine the directory and file name to get the full path
$fileScreensaverFullPath = Join-Path -Path $scriptDirectory -ChildPath $fileSaver


# Function to preload Chrome with two tabs
function Preload-Chrome {
    $chromeProcess = Get-Process -Name "chrome" -ErrorAction SilentlyContinue
    if (-not $chromeProcess) {
        Write-Host "Preloading Chrome with welcome.html in fullscreen..."
        $chromePath = "C:\Program Files\Google\Chrome\Application\chrome.exe"
        Start-Process $chromePath "--kiosk file:///$fullWelcome --new-window --start-fullscreen"
        Start-Sleep -Seconds 5
        Write-Host "Opening kiosk websites in Chrome..."
        Start-Process $chromePath "--kiosk $filePage" -NoNewWindow
        Start-Sleep -Seconds 1
        Start-Process $chromePath "--kiosk $fileScreensaverFullPath" -NoNewWindow
        Start-Sleep -Seconds 1
    } else {
        Write-Host "Chrome is already running."
    }
}

function Show-PageIntro {
    param (
        [string]$WebsiteName,
        [string]$KeySequence
    )

    Write-Host "Showing $WebsiteName website..."
    $chromeProcess = Get-Process -Name "chrome" -ErrorAction SilentlyContinue
    if ($chromeProcess) {
        $wshell = New-Object -ComObject wscript.shell
                Write-Host "Attempting to activate window with title: $($chromeProcess[0].MainWindowTitle)"

        $wshell.AppActivate($chromeProcess[0].MainWindowTitle)
        Start-Sleep -Milliseconds 500
        $wshell.SendKeys($KeySequence)
    }
}

function Show-Page {
    param (
        [string]$WebsiteName,
        [string]$KeySequence
    )

    Write-Host "Showing $WebsiteName website..."
    $chromeProcess = Get-Process -Name "chrome" -ErrorAction SilentlyContinue
    if ($chromeProcess) {
        $wshell = New-Object -ComObject wscript.shell
        Write-Host "Attempting to activate window with title: $($chromeProcess[0].MainWindowTitle)"
        $wshell.AppActivate($chromeProcess[0].MainWindowTitle)
        Start-Sleep -Milliseconds 500
        $wshell.SendKeys($KeySequence)
    }
}

# Function to mute the system volume
function Mute-Volume {
    Write-Host "Muting system volume..."
    Set-AudioDevice -PlaybackMute $true
}

# Function to unmute the system volume
function Unmute-Volume {
    Write-Host "Unmuting system volume..."
    Set-AudioDevice -PlaybackMute $false
}

# Function to deactivate the screensaver
function Deactivate-Screensaver {
   Write-Host "Deactivating screensaver..."

    # Unmute system volume
    Unmute-Volume
     # Wait a moment before showing the graphical intro
    Start-Sleep -Milliseconds 500
    $serialPort.Open()
    $serialPort.WriteLine("1")
    $serialPort.Close()
    Show-Page  -WebsiteName "Graphical Intro" -KeySequence "^1"
    Start-Sleep -Milliseconds 3000
    Show-Page  -WebsiteName "Kiosk" -KeySequence "^2"
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
        Write-Host "USB device count decreased. Triggering screensaver and sending 0 to serial port..."
        #Mute system Volume
        Mute-Volume
        Show-Page  -WebsiteName "Screensaver page" -KeySequence "^3"
        $serialPort.Open()
        $serialPort.WriteLine("0")
        $serialPort.Close()
    } elseif ($currentDeviceCount -gt $previousDeviceCount) {
        Write-Host "USB device count increased. Deactivating screensaver and sending 1 to serial port..."
        Deactivate-Screensaver
    }

    $previousDeviceCount = $currentDeviceCount
}

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

# Combine the directory and file name to get the full path
$fullPage = Join-Path -Path $scriptDirectory -ChildPath $fileLocal

# Ensure the InputSimulator type is defined
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class InputSimulator {
    [DllImport("user32.dll")]
    public static extern void mouse_event(int dwFlags, int dx, int dy, int dwData, UIntPtr dwExtraInfo);
    public const int MOUSEEVENTF_MOVE = 0x0001;

    [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
    public static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);
    public const int KEYEVENTF_EXTENDEDKEY = 0x1;
    public const int KEYEVENTF_KEYUP = 0x2;
    public const byte VK_SHIFT = 0x10;

    public static void MoveMouse() {
        mouse_event(MOUSEEVENTF_MOVE, 0, 1, 0, UIntPtr.Zero);
        mouse_event(MOUSEEVENTF_MOVE, 0, -1, 0, UIntPtr.Zero);
    }

    public static void PressKey() {
        keybd_event(VK_SHIFT, 0, KEYEVENTF_EXTENDEDKEY, UIntPtr.Zero);
        keybd_event(VK_SHIFT, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
    }
}
"@

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
        Start-Process $screensaverPath
    } else {
        Write-Host "Screensaver path not found: $screensaverPath"
    }
}

# Function to simulate mouse movement and key press
function Simulate-Input {
    Write-Host "Simulating mouse movement and key press..."
    [InputSimulator]::MoveMouse()
    Start-Sleep -Milliseconds 100
    [InputSimulator]::PressKey()
}

# Function to deactivate the screensaver
function Deactivate-Screensaver {
    Write-Host "Deactivating screensaver..."
    # Simulate input to ensure the screensaver exits
    Simulate-Input
    
    
    # Close the screensaver process if it's running
    $screensaverProcess = Get-Process -Name "VideoScreensaver" -ErrorAction SilentlyContinue
    if ($screensaverProcess) {
        Stop-Process -Name "VideoScreensaver" -Force
    }
    
    
    Show-GraphicalIntro
    Start-Sleep -Milliseconds 3000  # Wait a moment after showing the graphical intro
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
    Write-Host "Inside loop, waiting 100 ms..."
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


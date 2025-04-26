
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

# Define the game file name
$fileGame = "Echoia"

# Define the file name
$fileWelcome = "welcome.html"

# Define the file name
$fileSaver = "videosaver.html"


# Combine the directory and file name to get the full path
$fullWelcome = Join-Path -Path $scriptDirectory -ChildPath $fileWelcome


# Combine the directory and file name to get the full path
$fileScreensaverFullPath = Join-Path -Path $scriptDirectory -ChildPath $fileSaver


# Define the necessary user32.dll methods if they don't already exist
if (-not ([System.Management.Automation.PSTypeName]'User32').Type) {
    Add-Type @"
    using System;
    using System.Runtime.InteropServices;
    public class User32 {
        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool IsIconic(IntPtr hWnd);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll", SetLastError = true)]
        public static extern int GetWindowText(IntPtr hWnd, System.Text.StringBuilder lpString, int nMaxCount);

        public const int SW_RESTORE = 9;
    }
"@
}

# Define the necessary user32.dll methods if they don't already exist
if (-not ([System.Management.Automation.PSTypeName]'User32').Type) {
    Add-Type @"
    using System;
    using System.Runtime.InteropServices;
    public class User32 {
        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool BlockInput(bool fBlockIt);
    }
"@
}

# Function to block keyboard input
function Block-KeyboardInput {
    [User32]::BlockInput($true)
}

# Function to unblock keyboard input
function Unblock-KeyboardInput {
    [User32]::BlockInput($false)
}

# Function to focus a window by its process name
function Focus-WindowByProcessName {
    param (
        [string]$processName,
        [string]$windowTitle = $null
    )
    $found = $false
    try{
    $process = Get-Process -Name $processName -ErrorAction Stop
    foreach ($p in $process) {
        $hWnd = $p.MainWindowHandle
        if ($p.MainWindowHandle -ne [IntPtr]::Zero) {
                if ([User32]::IsIconic($p.MainWindowHandle)) {
                    [User32]::ShowWindowAsync($p.MainWindowHandle, [User32]::SW_RESTORE)
                }
                if ([User32]::SetForegroundWindow($p.MainWindowHandle)) {
                    return $true
                }
            }
    }
     # If no main window found, try by title
        if ($windowTitle) {
            $hWnd = [User32]::FindWindow($null, $windowTitle)
            if ($hWnd -ne [IntPtr]::Zero) {
                if ([User32]::IsIconic($hWnd)) {
                    [User32]::ShowWindowAsync($hWnd, [User32]::SW_RESTORE)
                }
                return [User32]::SetForegroundWindow($hWnd)
            }
        }
    } catch {
        Write-Host "Focus error: $_"
    }
}
    

# Function to preload Chrome
function Preload-Chrome {
    $chromeProcess = Get-Process -Name "chrome" -ErrorAction SilentlyContinue
    if (-not $chromeProcess) {
        Write-Host "Preloading Chrome with welcome.html in fullscreen..."
        $chromePath = "C:\Program Files\Google\Chrome\Application\chrome.exe"
        Start-Process $chromePath "--kiosk file:///$fullWelcome --new-window --start-fullscreen"
        Start-Sleep -Seconds 5
        Start-Process $chromePath "--kiosk $fileScreensaverFullPath" -NoNewWindow 
        Start-Sleep -Seconds 1
    } else {
        Write-Host "Chrome is already running."
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
        $wshell.AppActivate($chromeProcess[0].MainWindowTitle)
        Start-Sleep -Milliseconds 500
        $wshell.SendKeys($KeySequence)
    }
}

# Function to bring Unreal game back to focus
function Focus-UnrealGame {
    $attempts = 0
    $maxAttempts = 3
    $unrealFocused = $false
    $processNames = @("$fileGame", "$fileGame-Win64-Shipping")
    while (-not $unrealFocused -and $attempts -lt $maxAttempts) {
        foreach ($processName in $processNames) {
            $unrealFocused = Focus-WindowByProcessName -processName $processName -windowTitle "$fileGame"
            if ($unrealFocused) {
                Write-Host "Unreal game window found and activated."
                break
            } else {
                Write-Host "Failed to activate Unreal game window for process ${processName}. Retrying... ($($attempts + 1)/$maxAttempts)"
            }
        }
        Start-Sleep -Milliseconds 3000
        $attempts++
    }
}

# Function to deactivate the screensaver
function Deactivate-Screensaver {
   Write-Host "Deactivating screensaver..."
    
     # Wait a moment before showing the graphical intro
    Start-Sleep -Milliseconds 500
    $serialPort.Open()
    $serialPort.WriteLine("1")
    $serialPort.Close()
    Show-Page  -WebsiteName "Graphical Intro" -KeySequence "^1"
    Start-Sleep -Milliseconds 3000
    Focus-UnrealGame
}


# Function to count connected USB devices
function Count-USBDevices {
    Write-Host "Counting USB devices..."
    Get-WmiObject -Query "SELECT * FROM Win32_USBControllerDevice" | Measure-Object | Select-Object -ExpandProperty Count
}

# Main loop
Write-Host "Entering main loop..."
$previousDeviceCount = Count-USBDevices

Preload-Chrome

while ($true) {
    Write-Host "Inside loop, waiting 100 ms..."
    Start-Sleep -Milliseconds 100
    $currentDeviceCount = Count-USBDevices

    if ($currentDeviceCount -lt $previousDeviceCount) {
        Write-Host "USB device count decreased. Triggering screensaver and sending '0' to serial port..."
        Show-Page  -WebsiteName "Screensaver page" -KeySequence "^2"
        Block-KeyboardInput
        $serialPort.Open()
        $serialPort.WriteLine("0")
        $serialPort.Close()
    } elseif ($currentDeviceCount -gt $previousDeviceCount) {
        Write-Host "USB device count increased. Deactivating screensaver and sending '1' to serial port..."
         # Wait for 2 seconds before deactivating the screensaver
        Deactivate-Screensaver
        Unblock-KeyboardInput
    }
   
    $previousDeviceCount = $currentDeviceCount
}


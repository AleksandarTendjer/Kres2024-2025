
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
$fileSaver = "VideoScreensaver-1.0\VideoScreensaver.scr"

# Combine the directory and file name to get the full path
$fullWelcome = Join-Path -Path $scriptDirectory -ChildPath $fileWelcome

# Combine the directory and file name to get the full path
$fullSaver = Join-Path -Path $scriptDirectory -ChildPath $fileSaver



# Ensure the InputSimulator type is defined
if (-not ("InputSimulator" -as [Type])) {
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
}

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
        
    } else {
        Write-Host "Chrome is already running."
    }
}

# Function to show the intro screen by switching to the first tab in Chrome
function Show-GraphicalIntro {
    $serialPort.Open()
    $serialPort.WriteLine("1")
    $serialPort.Close()
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

# Function to trigger the screensaver
function Trigger-Screensaver {
   
    $chromeFocused = Focus-WindowByProcessName -processName "chrome"
    if ($chromeFocused) {
        Write-Host "Intro Chrome window found and activated."
    } else {
        Write-Host "Failed to activate Intro Chrome window."
    }

    $screensaverProcess = Get-Process -Name "VideoScreensaver" -ErrorAction SilentlyContinue
    if ($screensaverProcess) {
        Write-Host "Screensaver is already running. Skipping new instance,switching focus"
                $focused = Focus-WindowByProcessName -processName "VideoScreensaver" -windowTitle "VideoScreensaver"
        if (-not $focused) {
            Write-Host "Could not focus screensaver window, forcing new instance..."
            # If focus failed, kill existing instance and start new
            Stop-Process -Name "VideoScreensaver" -Force
            Start-Process $fullSaver
        }
        return
    }
    Write-Host "Triggering screensaver..."
    if (Test-Path $fullSaver) {
        # Start with WindowStyle Hidden to prevent flash
        Start-Process $fullSaver -WindowStyle Hidden
        # Give it time to launch before focusing
        Start-Sleep -Milliseconds 500
        Focus-WindowByProcessName -processName "VideoScreensaver" -windowTitle "VideoScreensaver"
    } else {
        Write-Host "Screensaver path not found: $fullSaver"
    }
}

# Function to simulate mouse movement and key press
function Simulate-Input {
    Write-Host "Simulating mouse movement and key press..."
    Start-Sleep -Milliseconds 100
    [InputSimulator]::PressKey()
}

# Function to deactivate the screensaver
function Deactivate-Screensaver {
    Write-Host "Deactivating screensaver..."
    # Simulate input to ensure the screensaver exits
    Simulate-Input
    Start-Sleep -Milliseconds 500
    
    
    # Close the screensaver process if it's running
     Get-Process -Name "VideoScreensaver" -ErrorAction SilentlyContinue | Stop-Process -Force

     # Wait a moment before showing the graphical intro
    Show-GraphicalIntro
    Start-Sleep -Milliseconds 1000
    # Wait a moment after showing the graphical introf
    Focus-UnrealGame
}

# Preload Chrome with two tabs
Preload-Chrome

# Function to count connected USB devices
function Count-USBDevices {
    Write-Host "Counting USB devices..."
    Get-WmiObject -Query "SELECT * FROM Win32_USBControllerDevice" | Measure-Object | Select-Object -ExpandProperty Count
}



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


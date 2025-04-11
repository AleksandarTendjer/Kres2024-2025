#Arduino alive
# Define the serial port settings
$portName = "COM9"  # Replace with your Arduino's COM port
$baudRate = 9600    # Match with the baud rate set in your Arduino sketch


# Create a .NET serial port object
$serialPort = New-Object System.IO.Ports.SerialPort
$serialPort.PortName = $portName
$serialPort.BaudRate = $baudRate


# This will give you the directory of the script
$scriptDirectory = $PSScriptRoot


# Define the file name
$fileWelcome = "welcome.html"


# Define the file name
$fileGame = "My project (2)"


# Define the file name
$fileSaver = "VideoScreensaver-1.0\VideoScreensaver.scr"


# Combine the directory and file name to get the full path
$fullWelcome = Join-Path -Path $scriptDirectory -ChildPath $fileWelcome


# Combine the directory and file name to get the full path
$fullSaver = Join-Path -Path $scriptDirectory -ChildPath $fileSaver






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


# Function to focus a window by its process name
function Focus-WindowByProcessName {
    param (
        [string]$processName,
        [string]$windowTitle = $null
    )


    $found = $false
    foreach ($p in Get-Process -Name $processName -ErrorAction SilentlyContinue) {
        $hWnd = $p.MainWindowHandle
        if ($hWnd -ne [IntPtr]::Zero) {
            Write-Host "Found window handle for process ${processName}: ${hWnd}"
            if ([User32]::IsIconic($hWnd)) {
                Write-Host "Window is minimized, restoring it..."
                [User32]::ShowWindowAsync($hWnd, 9) # 9 = SW_RESTORE
            }
            $result = [User32]::SetForegroundWindow($hWnd)
            if ($result) {
                Write-Host "Successfully set foreground window for ${processName}"
                $found = $true
            } else {
                Write-Host "Failed to set foreground window for ${processName}"
            }
            break
        } else {
            Write-Host "No window handle found for process ${processName}"
        }
    }
    
    # If not found by process name, try by window title
    if (-not $found -and $windowTitle) {
        Write-Host "Attempting to find window by title: ${windowTitle}"
        $hWnd = [User32]::FindWindow($null, $windowTitle)
        if ($hWnd -ne [IntPtr]::Zero) {
            Write-Host "Found window handle for title ${windowTitle}: ${hWnd}"
            if ([User32]::IsIconic($hWnd)) {
                Write-Host "Window is minimized, restoring it..."
                [User32]::ShowWindowAsync($hWnd, 9) # 9 = SW_RESTORE
            }
            $result = [User32]::SetForegroundWindow($hWnd)
            if ($result) {
                Write-Host "Successfully set foreground window for ${windowTitle}"
                $found = $true
            } else {
                Write-Host "Failed to set foreground window for ${windowTitle}"
            }
        } else {
            Write-Host "No window handle found for title ${windowTitle}"
        }
    }
    
    return $found
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




# Function to bring Unity game back to focus
function Focus-UnityGame {
    $attempts = 0
    $maxAttempts = 3
    $unityFocused = $false
    $processNames = @("$fileGame")
    while (-not $unityFocused -and $attempts -lt $maxAttempts) {
        foreach ($processName in $processNames) {
            $unityFocused = Focus-WindowByProcessName -processName $processName -windowTitle "$fileGame"
            if ($unityFocused) {
                Write-Host "Unity game window found and activated."
                break
            } else {
                Write-Host "Failed to activate Unity game window for process ${processName}. Retrying... ($($attempts + 1)/$maxAttempts)"
            }
        }
        Start-Sleep -Milliseconds 500
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
    Start-Sleep -Milliseconds 150
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
    
     # Wait a moment before showing the graphical intro
    
    Show-GraphicalIntro
    Start-Sleep -Milliseconds 3000
     # Wait a moment after showing the graphical introf
    Focus-UnityGame
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
        $serialPort.Open()
        $serialPort.WriteLine("0")
        $serialPort.Close()
    } elseif ($currentDeviceCount -gt $previousDeviceCount) {
        Write-Host "USB device count increased. Deactivating screensaver and sending '1' to serial port..."
         # Wait for 2 seconds before deactivating the screensaver
        Deactivate-Screensaver
    }


    $previousDeviceCount = $currentDeviceCount
}



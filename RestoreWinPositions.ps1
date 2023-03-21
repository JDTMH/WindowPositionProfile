<#
The purpose of this script is to allow quick and easy restoration of saved window size and position.
Author:JDTMH
Date: 03/21/2023
To do: 
#>



#Reads json file into a configuration object 
Function ParseConfig($configpath)
{
  $jsontmp = Get-Content $configpath | Out-String | ConvertFrom-Json
  $global:config = {$jsontmp}.Invoke()
}

#Modifies a given window to a given size and position using the Windows API
Function MoveWindow($windowjson)
{
  $window = $windowjson | ConvertFrom-Json
  # Extract the name from the handle object using regex
  $windowName = $handle.Name -replace '^.*?\((.*?)\)$', '$1'
  Write-Host "Searching for program's with the following name: $windowName"
  $handle = Get-Process | Where-Object {$_.MainWindowTitle -like "*$($window.windowname)*"} | Select-Object -First 1

  #Select the program instance with a valid window 
  foreach ($h in $handle)
  {
      if ($h.MainWindowHandle -ne 0) {
        $takeHandle = $h.MainWindowHandle
        break
      }
  }

  #Skip if the program is not currently running
 if($takeHandle) {
    # Extract the name from the handle object using regex
    $windowName = $handle.Name -replace '^.*?\((.*?)\)$', '$1'
    Write-Host "Moving Program $windowName"
    $return = [Window]::MoveWindow($takeHandle, $window.x, $window.y, $window.width, $window.height, $True)
    Write-Host "Success?" $return "`n"
  }
  else {
    Write-Host "Program $windowName not running!"
  }
}

# =======
#  MAIN
# =======

#Window object type definition to interact with the Windows API
Try{
      [void][Window]
} 
Catch {
  Add-Type @"
using System;
using System.Runtime.InteropServices;
public class Window {
  [DllImport("user32.dll")]
  [return: MarshalAs(UnmanagedType.Bool)]
  public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
  [DllImport("User32.dll")]
  public extern static bool MoveWindow(IntPtr handle, int x, int y, int width, int height, bool redraw);
}
public struct RECT
{
  public int Left;        // x position of upper-left corner
  public int Top;         // y position of upper-left corner
  public int Right;       // x position of lower-right corner
  public int Bottom;      // y position of lower-right corner
}
"@
}

$global:config = New-Object System.Collections.ArrayList
#Reading file with configured size and position values
#Use `SaveWinPositions.ps1` to record these values
$configpath = $env:USERPROFILE + "\windowlayout.config"

#Modify windows' size and position as spec'd in the configuration file
ParseConfig($configpath)
$global:config | ForEach-Object { MoveWindow($_) }
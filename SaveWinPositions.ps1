<#
The purpose of this script is to allow quick and easy recording of current window size and position.
Author:JDTMH
Date: 03/21/2023
To do: 
#>


#Saves configuration object data to a json file
Function SaveConfig($configpath)
{
  $global:config | ConvertTo-Json | Out-File $configpath
}
#Queries the Windows API for a given window's size and position
#Saves it to a configuration object
Function GetWindowData($windowName)
{
  $matchingWindows = Get-Process | Where-Object {$_.MainWindowTitle -like "*$windowName*"}
if ($matchingWindows.Count -gt 0) {
    $takeHandle = $matchingWindows[0].MainWindowHandle
}
else {
    Write-Host "Program matching '$windowName' not found"
    return
}

  #Skip if the program is not found
  if($takeHandle -ne 0) {
    Write-Host "Recording data for window $windowName-$takeHandle ..."

    # This is to get the current window state
    $Rectangle = New-Object RECT
    $Return = [Window]::GetWindowRect($takeHandle,[ref]$Rectangle)
    $windowObject = New-Object -TypeName psobject
    $windowObject | Add-Member -MemberType NoteProperty -Name windowname -Value $windowName
    $windowObject | Add-Member -MemberType NoteProperty -Name x -Value $Rectangle.Left
    $windowObject | Add-Member -MemberType NoteProperty -Name y -Value $Rectangle.Top
    $windowObject | Add-Member -MemberType NoteProperty -Name width -Value ($Rectangle.Right - $Rectangle.Left)
    $windowObject | Add-Member -MemberType NoteProperty -Name height -Value ($Rectangle.Bottom - $Rectangle.Top)
    $json = $windowObject | ConvertTo-Json
    $global:config.Add($json)
  }
  else {
    Write-Host "Window $windowName not found"
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
$configpath = $env:USERPROFILE + "\windowlayout.config"

#Creating an array with the window names to record size and position
$windows = "Teams","Outlook","Vivaldi","Notepad++","OneNote","Excel","Word","Spotify"

#Read and record window sizes and positions
$windows | ForEach-Object { GetWindowData($_) }
SaveConfig($configpath)
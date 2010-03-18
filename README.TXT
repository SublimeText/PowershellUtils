================================================================================
RunExternalPSCommand Plugin for the Sublime Text text editor
================================================================================

This plugin provides an interface to filter text through a Windows
Powershell (PoSh) pipeline.

Requirements:
    * Windows Powershell v2. Windows Powershell v2 is preinstalled in
      Windows 7 and later and it's available for previous versions of
      Windows too.

Usage:
    * Assign a keybinding to "runExternalPSCommand".
    * Type in your PoSh command (or intrinsic, see below) and press the enter
      key.
    * Sublime Text selection's text is available to PoSh in $args[0] and, for
      convenience, in $a too.

Intrinsic commands:
    * !h will bring up the session's history of successful commands.
    * !mkh will persist the current history between sessions.

Examples:
    * $a -replace "\\","/"
    * "$(date)"
    * "$pwd"
    * [environment]::GetFolderPath([environment+specialfolder]::MyDocuments)
    * "!"*80; 1..100 | %{"Sublime Text rocks!"}; "!"*80
    * "ПАССАЖИРАМ mañana Schloß`n"*2 # (for testing)

Caveats:
    * Filtering many regions is sloooooooooooooooooooow.
    * To start a PoSh subshell, do either `Start-Process powershell` or
      `start powershell` (cmd.exe's start program), but don't call PoSh
      directly because it will be launched in windowless mode and will block
      Sublime Text forever. Should this happen to you, you can execute the
      following PoSh command from an actual PoSh prompt to terminate all PoSh
      processes except for the current session:

      Get-Process powershell | Where-Object { $_.Id -ne $PID } | Stop-Process

      ...or the equivalent, shorter version:

      gps powershell|?{ $_.id -ne $pid }|stop-process # compact version

      TODO: This issue should probably be fixed.

Bugs:
    * Start Powershell, start ST from poweshell, use this plugin.
      You get a Windows invalid handle error.
      Start-Process ST works...?
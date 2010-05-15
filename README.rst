
PowershellUtils package for Sublime Text
========================================

This plugin provides an interface to filter text through a Windows Powershell pipeline.

Requirements
************

* ``Windows Powershell v2.``
  Windows Powershell v2 is preinstalled in Windows 7 and later and it's available for previous versions of Windows too.

Side effects
************

This plugin shouldn't have any side effects unless you execute it explicitly. Particulartly, it doesn't define any keybindings or react to events.

Bear in mind, though, that Windows Powershell is a powerful languange and needs to be used carefully to avoid undesired effects.

Getting started
***************

You need to define a keybinding for the command ``runExternalPSCommand`` or run it from the console like so: ``view.runCommand("runExternalPSCommand")``.

Usage
*****

Using The Windows Powershell Pipeline
-------------------------------------

1. Execute ``runExternalPSCommand``
2. Type in your Windows Powershell command
3. Press the enter key

All the currently selected regions in Sublime Text will be piped into your command. You can access each of this regions in turn through the ``$_`` automatic variable.

Roughly, this is what goes on behind the scenes::

    reg1..regN | <your command> | out-string

You can ignore the piped content and treat your command as the start point of the pipeline.

The generated output will be inserted into each region in turn.

Examples
********

``$_.toupper()``
    Turns each region's content to all caps.
``$_ -replace "\\","/"``
    Replaces each region's content as indicated.
``"$(date)"``
    Replaces each region's content with the current date.
``"$pwd"``
    Replaces each region's content with the current working directory.
``[environment]::GetFolderPath([environment+specialfolder]::MyDocuments)``
    Replaces each region's content with the path to the user's ``My Documents`` folder.
``0..6|%{ "$($_+1) $([dayofweek]$_)" }``
    Replaces each region's content with the enumerated week days.

Caveats
*******

To start a Windows Powershell shell, do either ``Start-Process powershell`` or ``cmd /k start powershell``, but don't call Windows Powershell directly because it will be launched in windowless mode and will block Sublime Text forever. Should this happen to you, you can execute the following command from an actual Windows Powershell prompt to terminate all Windows Powershell processes except for the current session::

    Get-Process powershell | Where-Object { $_.Id -ne $PID } | Stop-Process

Alternatively, you can use a shorter version::

    gps powershell|?{$_.id -ne $pid}|kill

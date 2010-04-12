# coding=utf-8
from __future__ import with_statement
import sublime, sublimeplugin
import os.path
import subprocess
import codecs
import base64
import ctypes

PATH_TO_THIS_FILE = os.path.dirname(os.path.abspath(__file__))

def getPathToPoShScript():
    return os.path.join(PATH_TO_THIS_FILE, "psbuff.ps1")

def getPathToPoShHistoryDB():
    return os.path.join(PATH_TO_THIS_FILE, "pshist.txt")

# The PoSh pipeline provided by the user is merged with this template and the
# resulting PoSh script is passed the text cotained in the selected region in
# Sublime Text. If many regions exist, they are filtered one after the other.
#
# NOTE: PoSh accepts a command as a base64 encoded string, but that way
# we'd fill up the Windows console's buffer quicker and besides
# any error info will (apparently) be returned as an XML string
PoSh_SCRIPT_TEMPLATE = """
$a = $args[0]
[void] $(chcp 65001) # we want utf-8 returned from the console!
# We receive a base64 encoded UTF16LE encoding from the command line.
$args[0] = ($a = [text.encoding]::Unicode.getstring([convert]::Frombase64String($a)))
# +++ Lines up to here inserted by ExecutePSCommand plugin for Sublime Text +++
%s
# +++ Lines from here inserted by ExecutePSCommand plugin Sublime Text+++
""".decode('utf-8')

class RunExternalPSCommandCommand(sublimeplugin.TextCommand):
    """
    This plugin provides an interface to filter text through a Windows
    Powershell (PoSh) pipeline. See README.TXT for instructions.
    """

    PoSh_HISTORY_MAX_LENGTH = 50

    def __init__(self, *args, **kwargs):
        self.PSHistory = getPoShSavedHistory()
        self.lastFailedCommand = ""
        super(sublimeplugin.TextCommand, self).__init__(*args, **kwargs)

    def _addToPSHistory(self, command):
        if not command in self.PSHistory:
            self.PSHistory.insert(0, command)
        if len(self.PSHistory) > self.PoSh_HISTORY_MAX_LENGTH:
            self.PSHistory.pop()

    def _showPSHistory(self, view):
        view.window().showQuickPanel('', "runExternalPSCommand", self.PSHistory,
                                        sublime.QUICK_PANEL_MONOSPACE_FONT)

    def _parseIntrinsicCommands(self, userPoShCmd, view):
        if userPoShCmd == '!h':
            if self.PSHistory:
                self._showPSHistory(view)
            else:
                sublime.statusMessage("Powershell command history is empty.")
            return True
        if userPoShCmd == '!mkh':
            try:
                with open(getPathToPoShHistoryDB(), 'w') as f:
                    cmds = [(cmd + '\n').encode('utf-8') for cmd in self.PSHistory]
                    f.writelines(cmds)
                    sublime.statusMessage("Powershell command history saved.")
                return True
            except IOError:
                sublime.statusMessage("ERROR: Could not save Powershell command history.")
        else:
            return False

    def run(self, view, args):
        def onDone(userPoShCmd):
            if self._parseIntrinsicCommands(userPoShCmd, view): return

            # UTF8 signature required! If you don't use it, the sky will fall!
            # The Windows console won't interpret correctly a UTF8 encoding
            # without a signature.
            try:
                with codecs.open(getPathToPoShScript(), 'w', 'utf_8_sig') as f:
                    f.write( (PoSh_SCRIPT_TEMPLATE % userPoShCmd) )
            except IOError:
                sublime.statusMessage("ERROR: Could not access Powershell script file.")

            for region in view.sel():
                try:
                    PoShOutput, PoShErrInfo = filterThruPoSh(view.substr(region))
                # Catches errors for any OS, not just Windows.
                except EnvironmentError, e:
                    # TODO: This catches too many errors?
                    sublime.errorMessage("Windows error. Possible causes:\n\n" +
                                          "* Is Powershell in your %PATH%?\n" +
                                          "* Use Start-Process to start ST from Powershell.\n\n%s" % e)
                    return

                if PoShErrInfo:
                    print PoShErrInfo
                    sublime.statusMessage("PowerShell error.")
                    view.window().runCommand("showPanel console")
                    self.lastFailedCommand = userPoShCmd
                    return
                elif PoShOutput:
                    self.lastFailedCommand = ''
                    # NOTE 1: If you call Popen with shell=True and the length of PoShOutput
                    # exceeds the console's buffer width, the output will be split with
                    # extra, unexpected \r\n at the corresponding spots.
                    # NOTE 2: PoSh can return XML too if you need it.
                    self._addToPSHistory(userPoShCmd)
                    # PS will insert \r\n at the end of every line--normalize.
                    view.replace(region, PoShOutput[:-2].replace('\r\n', '\n'))

        initialText = args[0] if args else self.lastFailedCommand
        view.window().showInputPanel("PoSh cmd:", initialText, onDone, None, None)

def getPoShSavedHistory():
    # If the command history file doesn't exist now, it will be created when
    # the user chooses to persist the current history for the first time.
    try:
        with open(getPathToPoShHistoryDB(), 'r') as f:
            return [command[:-1].decode('utf-8') for command in f.readlines()]
    except IOError:
        return []

def getConsoleCodePage():
    # Retrieve codepage with Win API call and use that.
    # TODO: This doesn't work...
    # codepage = "cp%s" % str(ctypes.windll.kernel32.GetConsoleCP())
    codepage = subprocess.Popen(["chcp"], shell=True, stdout=subprocess.PIPE).communicate()[0]
    codepage = "cp" + codepage[:-2].split(" ")[-1:][0].strip()
    return codepage

def buildPoShCmdLine(pathToScriptFile, argsToScript=""):
    return ["powershell",
                        "-noprofile",
                        "-nologo",
                        "-noninteractive",
                        # PoSh 2.0 lets you specify an ExecutionPolicy
                        # from the cmdline, but 1.0 doesn't.
                        "-executionpolicy", "remotesigned",
                        "-file", pathToScriptFile,
                        # According to Popen, CreateProcess doesn't allow strings containing
                        # nulls as parameters. Besides, PoSh will get confused if we pass
                        # UTF8 strings as args to the script. Do the following instead:
                        #   * Encode parameter string in UTF16LE (NET Framework likes that).
                        #   * Encode resulting bytestring in base64.
                        #   * Decode in the PoSh script.
                        base64.b64encode(argsToScript.encode("utf-16LE")),]

def filterThruPoSh(text):
    # Hide the child process window.
    startupinfo = subprocess.STARTUPINFO()
    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW

    PoShOutput, PoShErrInfo = subprocess.Popen(buildPoShCmdLine(getPathToPoShScript(), text),
                                            shell=False, # TODO: Needed?
                                            stdout=subprocess.PIPE,
                                            stderr=subprocess.PIPE,
                                            startupinfo=startupinfo).communicate()

    # We've changed the Windows console's default codepage in the PoSh script
    # by calling chcp 65001. Therefore, now we need to decode a UTF8 stream
    # with sinature.
    # Note: PoShErrInfo still gets encoded in the default codepage.
    return ( PoShOutput.decode('utf_8_sig'),
             PoShErrInfo.decode(getConsoleCodePage()), )


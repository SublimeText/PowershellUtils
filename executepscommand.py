from __future__ import with_statement
import sublime, sublimeplugin
import os.path
import subprocess
import codecs
import base64
import ctypes
import glob

joinToThisFileParent = lambda fileName: os.path.join(
                                    os.path.dirname(os.path.abspath(__file__)),
                                    fileName
                                    )

def dumpRegions(rgs):
    """Saves regions to disk."""
    for f in glob.glob(joinToThisFileParent("tmp/*.txt")):
        os.remove(f)

    try:
        for i, r in enumerate(rgs):
            f = open(joinToThisFileParent("tmp/in%d.txt" % i), "w")
            f.write(r.encode("utf_8_sig"))
            f.close()
    except TypeError:
        f = open(joinToThisFileParent("tmp/in0.txt"), "w")
        f.write(rgs.encode("utf_8_sig"))
        f.close()

def getOutputs():
    for f in sorted(glob.glob(joinToThisFileParent("tmp/out*.txt"))):
        yield open(f, "r").read().decode("utf8")[:-1]

def getPathToPoShScript():
    return joinToThisFileParent("psbuff.ps1")

def getPathToPoShHistoryDB():
    return joinToThisFileParent("pshist.txt")

# The PoSh pipeline provided by the user is merged with this template.
PoSh_SCRIPT_TEMPLATE = u"""
$script:i = 0
get-item "$(split-path $MyInvocation.mycommand.path -parent)\\tmp\\in*.txt" | `
    foreach-object {
        $a = get-content -path $_
        %s | out-file `
                        "$(split-path $MyInvocation.mycommand.path -parent)\\tmp\\out${script:i}.txt" `
                        -append `
                        -encoding utf8 `
                        -force
        ++$script:i
    }
"""


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

            dumpRegions(view.substr(r) for r in view.sel())

            # UTF8 signature required! If you don't use it, the sky will fall!
            # The Windows console won't interpret correctly a UTF8 encoding
            # without a signature.
            try:
                with codecs.open(getPathToPoShScript(), 'w', 'utf_8_sig') as f:
                    f.write( PoSh_SCRIPT_TEMPLATE % (userPoShCmd) )
            except IOError:
                sublime.statusMessage("ERROR: Could not access Powershell script file.")
                return

            try:
                PoShOutput, PoShErrInfo = filterThruPoSh("")
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
            else:
                self.lastFailedCommand = ''
                self._addToPSHistory(userPoShCmd)
                # cannot do zip(regs, outputs) because view.sel() maintains
                # regions up-to-date if any of them changes.
                for i, txt in enumerate(getOutputs()):
                    view.replace(view.sel()[i], txt)

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

def getOEMCP():
    # Windows OEM/Ansi codepage mismatch issue.
    # We need the OEM cp, because powershell is a console program.
    codepage = ctypes.windll.kernel32.GetOEMCP()
    return str(codepage)

def buildPoShCmdLine(pathToScriptFile, argsToScript=""):
    return ["powershell",
                        "-noprofile",
                        "-nologo",
                        "-noninteractive",
                        # PoSh 2.0 lets you specify an ExecutionPolicy
                        # from the cmdline, but 1.0 doesn't.
                        "-executionpolicy", "remotesigned",
                        "-file", pathToScriptFile, ]

def filterThruPoSh(text):
    # Hide the child process window.
    startupinfo = subprocess.STARTUPINFO()
    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW

    PoShOutput, PoShErrInfo = subprocess.Popen(buildPoShCmdLine(getPathToPoShScript(), text),
                                            stdout=subprocess.PIPE,
                                            stderr=subprocess.PIPE,
                                            startupinfo=startupinfo).communicate()

    # We've changed the Windows console's default codepage in the PoSh script
    # Therefore, now we need to decode a UTF8 stream with sinature.
    # Note: PoShErrInfo still gets encoded in the default codepage.
    return ( PoShOutput.decode("cp850"),
             PoShErrInfo.decode(getOEMCP()), )

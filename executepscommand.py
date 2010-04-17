from __future__ import with_statement
import sublime, sublimeplugin
import os.path
import subprocess
import codecs
import ctypes
from xml.etree.ElementTree import ElementTree

# Things to remember:
#   * encoding expected by Windows program (console/gui).
#   * encoding expected by Powershell.

# The PoSh pipeline provided by the user and the input values (regions)
# are merged with this template.
PoSh_SCRIPT_TEMPLATE = u"""
$script:pathToOutPutFile ="$(split-path $MyInvocation.mycommand.path -parent)\\tmp\\out.txt"
"<outputs>" | out-file $pathToOutPutFile -encoding utf8 -force
$script:regionTexts = %s
$script:regionTexts | foreach-object {
                        %s | foreach-object { "<out><![CDATA[$_]]></out>`n" } | out-file `
                                                                        -filepath $pathToOutPutFile `
                                                                        -append `
                                                                        -encoding utf8
}
"</outputs>" | out-file $pathToOutPutFile -encoding utf8 -append -force
"""

class CantAccessScriptFileError(Exception):
    pass

joinToThisFileParent = lambda fileName: os.path.join(
                                    os.path.dirname(os.path.abspath(__file__)),
                                    fileName
                                    )

def regionsToPoShArray(view, rgs):
    return ",".join("'%s'" % view.substr(r).replace("'", "''") for r in rgs)

def getOutputs():
    tree = ElementTree()
    tree.parse(joinToThisFileParent("tmp/out.txt"))
    return [el.text for el in tree.findall("out")]

def getPathToPoShScript():
    return joinToThisFileParent("psbuff.ps1")

def getPathToPoShHistoryDB():
    return joinToThisFileParent("pshist.txt")


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
            # User doesn't want to filter anything.
            if self._parseIntrinsicCommands(userPoShCmd, view): return

            try:
                PoShOutput, PoShErrInfo = filterThruPoSh(regionsToPoShArray(view, view.sel()), userPoShCmd)
            except EnvironmentError, e:
                sublime.errorMessage("Windows error. Possible causes:\n\n" +
                                      "* Is Powershell in your %PATH%?\n" +
                                      "* Use Start-Process to start ST from Powershell.\n\n%s" % e)
                return
            except CantAccessScriptFileError:
                sublime.errorMessage("Cannot access script file.")
                return

            # Inform the user that something went wrong in his PoSh code or
            # perform substitutions and do house-keeping.
            if PoShErrInfo:
                print PoShErrInfo
                sublime.statusMessage("PowerShell error.")
                view.window().runCommand("showPanel console")
                self.lastFailedCommand = userPoShCmd
                return
            else:
                self.lastFailedCommand = ''
                self._addToPSHistory(userPoShCmd)
                # Cannot do zip(regs, outputs) because view.sel() maintains
                # regions up-to-date if any of them changes.
                for i, txt in enumerate(getOutputs()):
                    view.replace(view.sel()[i], txt)

        # Open cmd line.
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

def buildScript(values, userPoShCmd):
    with codecs.open(getPathToPoShScript(), 'w', 'utf_8_sig') as f:
        f.write( PoSh_SCRIPT_TEMPLATE % (values, userPoShCmd) )

def buildPoShCmdLine():
    return ["powershell",
                        "-noprofile",
                        "-nologo",
                        "-noninteractive",
                        # PoSh 2.0 lets you specify an ExecutionPolicy
                        # from the cmdline, but 1.0 doesn't.
                        "-executionpolicy", "remotesigned",
                        "-file", getPathToPoShScript(), ]

def filterThruPoSh(values, userPoShCmd):

    try:
        buildScript(values, userPoShCmd)
    except IOError:
        raise CantAccessScriptFileError

    # Hide the child process window.
    startupinfo = subprocess.STARTUPINFO()
    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW

    PoShOutput, PoShErrInfo = subprocess.Popen(buildPoShCmdLine(),
                                            stdout=subprocess.PIPE,
                                            stderr=subprocess.PIPE,
                                            startupinfo=startupinfo).communicate()

    return ( PoShOutput.decode(getOEMCP()),
             PoShErrInfo.decode(getOEMCP()), )

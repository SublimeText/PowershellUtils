from __future__ import with_statement
import sublime, sublimeplugin
import os.path
import subprocess
import codecs
import ctypes
import tempfile
import sublimepath
import functools
from xml.etree.ElementTree import ElementTree

# The PoSh pipeline provided by the user and the input values (regions)
# are merged with this template.
PoSh_SCRIPT_TEMPLATE = u"""
function collectData { "<out><![CDATA[$([string]::join('`n', $input))]]></out>`n" }
$script:pathToOutPutFile ="%s"
"<outputs>" | out-file $pathToOutPutFile -encoding utf8 -force
$script:regionTexts = %s
$script:regionTexts | foreach-object {
                        %s | out-string | collectData | out-file `
                                                    -filepath $pathToOutPutFile `
                                                    -append `
                                                    -encoding utf8
}
"</outputs>" | out-file $pathToOutPutFile -encoding utf8 -append -force
"""

THIS_PACKAGE_NAME = "PowershellUtils"
THIS_PACKAGE_DEV_NAME = "XXX" + THIS_PACKAGE_NAME
POSH_SCRIPT_FILE_NAME = "psbuff.ps1"
POSH_HISTORY_DB_NAME = "pshist.txt"
OUTPUT_SINK_NAME = "out.xml"
DEBUG = os.path.exists(sublime.packagesPath() + "/" + THIS_PACKAGE_DEV_NAME)


class CantAccessScriptFileError(Exception):
    pass


def regionsToPoShArray(view, rgs):
    """
    Return a PoSh array: 'x', 'y', 'z' ... and escape single quotes like
    this : 'escaped ''sinqle quoted text'''
    """
    return ",".join("'%s'" % view.substr(r).replace("'", "''") for r in rgs)

def getOutputs():
    tree = ElementTree()
    tree.parse(getPathToOutputSink())
    return [el.text[:-1] for el in tree.findall("out")]

def getThisPackageName():
    """
    Name varies depending on the name of the folder containing this code.
    TODO: Is __name__ accurate in Sublime? __file__ doesn't seem to be.
    """
    return THIS_PACKAGE_NAME if not DEBUG else THIS_PACKAGE_DEV_NAME

def getPathToPoShScript():
    return sublimepath.rootAtPackagesDir(getThisPackageName(), POSH_SCRIPT_FILE_NAME)

def getPathToPoShHistoryDB():
    return sublimepath.rootAtPackagesDir(getThisPackageName(), POSH_HISTORY_DB_NAME)

def getPathToOutputSink():
    return sublimepath.rootAtPackagesDir(getThisPackageName(), OUTPUT_SINK_NAME)

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
        f.write( PoSh_SCRIPT_TEMPLATE % (getPathToOutputSink(), values, userPoShCmd) )

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

        if args and len(args) > 1:
            self.onDone(view, args[1])
            return

        # Open cmd line.
        initialText = args[0] if args else self.lastFailedCommand
        inputPanel = view.window().showInputPanel("PoSh cmd:", initialText, functools.partial(self.onDone, view), None, None)

    def onDone(self, view, userPoShCmd):
        # Exit if user doesn't actually want to filter anything.
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

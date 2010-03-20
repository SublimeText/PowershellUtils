# coding=utf-8
from __future__ import with_statement
import sublime, sublimeplugin
import os.path
import subprocess
import codecs
import base64

def getPathToPoShScript():
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(BASE_DIR, "psbuff.ps1")
    # PS_HISTORY_FILE = os.path.join(BASE_DIR, "pshist.txt")

# The PoSh pipeline provided by the user is merged with this template and the
# resulting PoSh script is passed the text cotained in the selected region in
# Sublime Text. If many regions exist, they are filtered one after the other.
PS_SCRIPT_TEMPLATE = """
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

    PS_HISTORY_MAX_LENGTH = 50
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    PS_BUFF_FILE = getPathToPoShScript()
    PS_HISTORY_FILE = os.path.join(BASE_DIR, "pshist.txt")

    def __init__(self, *args, **kwargs):
        self.PSHistory = []
        self.lastFailedCommand = ''

        # If the command history file doesn't exist now, it will be created when
        # the user chooses to persist the current history for the first time.
        try:
            with open(self.PS_HISTORY_FILE, 'r') as f:
                self.PSHistory = [command[:-1] for command in f.readlines()]
        except IOError:
            pass

        super(sublimeplugin.TextCommand, self).__init__(*args, **kwargs)

    def addToPSHistory(self, command):
        if not command in self.PSHistory:
            self.PSHistory.insert(0, command)
        if len(self.PSHistory) > self.PS_HISTORY_MAX_LENGTH:
            self.PSHistory.pop()

    def showPSHistory(self, view):
        view.window().showQuickPanel('', "runExternalPSCommand", self.PSHistory,
                                        sublime.QUICK_PANEL_MONOSPACE_FONT)

    def run(self, view, args):
        def onDone(PSCommand):
            # Intrinsic commands.
            if PSCommand == '!h':
                if self.PSHistory:
                    self.showPSHistory(view)
                else:
                    sublime.statusMessage("PoSh command history is empty.")
                return
            if PSCommand == '!mkh':
                with open(self.PS_HISTORY_FILE, 'w') as f:
                    f.writelines([(x + '\n') for x in self.PSHistory])
                    sublime.statusMessage("PoSh command history saved.")
                return

            # PoSh accepts a command as a base64 encoded string, but that way
            # we'd fill up the Windows console's buffer quicker and besides
            # any error info will (apparently) be returned as an XML string
            PSScriptContent = (PS_SCRIPT_TEMPLATE % PSCommand)

            # UTF8 signature required! If you don't use it, the sky will fall!
            # The Windows console won't interpret correctly a UTF8 encoding
            # without a signature.
            with codecs.open(self.PS_BUFF_FILE, 'w', 'utf_8_sig') as f:
                f.write(PSScriptContent)

            for region in view.sel():
                text = view.substr(region)

                try:
                    startupinfo = subprocess.STARTUPINFO()
                    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                    PSOutput, PSErrInfo = subprocess.Popen(getPoShCmdLine(self.PS_BUFF_FILE, text),
                                                            shell=False, # TODO: Needed?
                                                            stdout=subprocess.PIPE,
                                                            stderr=subprocess.PIPE,
                                                            startupinfo=startupinfo).communicate()

                    # We've changed the Windows console's default codepage in the PoSh script
                    # by calling chcp 65001. Therefore, now we need to decode a UTF8 stream
                    # with sinature.
                    # Note: PSErrInfo still gets encoded in the default codepage.
                    PSOutput, PSErrInfo = (PSOutput.decode('utf_8_sig'),
                                          PSErrInfo.decode(getDOSPromptDefaultCodepage()))

                # Catches errors for any OS, not just Windows.
                except EnvironmentError, e:
                    # TODO: This catches too many errors?
                    sublime.errorMessage("Windows error. Possible causes:\n\n" +
                                          "* Is Powershell in your %PATH%?\n" +
                                          "* Use Start-Process to start ST from PoSh.\n\n%s" % e)
                    return

                if PSErrInfo:
                    print PSErrInfo
                    sublime.statusMessage("PowerShell error.")
                    view.window().runCommand("showPanel console")
                    self.lastFailedCommand = PSCommand
                    return
                elif PSOutput:
                    self.lastFailedCommand = ''
                    # NOTE 1: If you call Popen with shell=True and the length of PSOutput
                    # exceeds the console's buffer width, the output will be split with
                    # extra, unexpected \r\n at the corresponding spots.
                    # NOTE 2: PoSh can return XML too if you need it.
                    self.addToPSHistory(PSCommand)
                    # PS will insert \r\n at the end of every line--normalize.
                    view.replace(region, PSOutput[:-2].replace('\r\n', '\n'))

        # Don't make the user retype the last unsuccessful command
        # if it was a Powershell error.
        initialText = args[0] if args else self.lastFailedCommand
        w = view.window()
        w.showInputPanel("PS cmd:", initialText, onDone, None, None)

def getDOSPromptDefaultCodepage():
    # Retrieve codepage with Win API call and use that.
    codepage = subprocess.Popen(["chcp"], shell=True, stdout=subprocess.PIPE).communicate()[0]
    codepage = "cp" + codepage[:-2].split(" ")[-1:][0].strip()
    return codepage

def getPoShCmdLine(script, argsToScript=""):
    return ["powershell",
                        "-noprofile",
                        "-nologo",
                        "-noninteractive",
                        # PoSh 2.0 lets you specify an ExecutionPolicy
                        # from the cmdline, but 1.0 doesn't.
                        "-executionpolicy", "remotesigned",
                        "-file", script,
                        # According to Popen, CreateProcess doesn't allow strings containing
                        # nulls as parameters. Besides, PoSh will get confused if we pass
                        # UTF8 strings as args to the script. Do the following instead:
                        #   * Encode parameter string in UTF16LE (NET Framework likes that).
                        #   * Encode resulting bytestring in base64.
                        #   * Decode in the PoSh script.
                        base64.b64encode(argsToScript.encode("utf-16LE")),]


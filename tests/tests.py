import unittest
import sublimeunittest
#===============================================================================
#   Import your plugins and any Python lib available on your system.
#   Add your tests below this header.
#   Remember you're using your system's Python!
#===============================================================================

import sublime
import sublimeplugin
import executepscommand
import ctypes

class HelpersTestCase(unittest.TestCase):
    def setUp(self):
        self.v = sublime.View()

    def testReturnCorrectPoShCmdLine(self):
        cmdLine = executepscommand.buildPoShCmdLine()
        expected = ["powershell",
                        "-noprofile",
                        "-nologo",
                        "-noninteractive",
                        # PoSh 2.0 lets you specify an ExecutionPolicy
                        # from the cmdline, but 1.0 doesn't.
                        "-executionpolicy", "remotesigned",
                        "-file" ] # sublimemocks returns weird path
        self.assertEquals(cmdLine[:-1], expected)

    def testReturnCorrectOEMCopePage(self):
        actual = executepscommand.getOEMCP()
        expected = str(ctypes.windll.kernel32.GetOEMCP())
        self.assertEquals(actual, expected)


if __name__ == "__main__":
    unittest.main()

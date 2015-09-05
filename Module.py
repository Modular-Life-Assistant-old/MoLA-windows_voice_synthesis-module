from helpers.modules.VoiceSynthesisModule import VoiceSynthesisModule

import os
try:
    import pythoncom
    import win32com.client
except ImportError:
    if os.name == 'nt':
        from core import Log
        Log.crash('python package "pywin32" is required')


class Module(VoiceSynthesisModule):
    voice_quality = 10

    def is_available(self):
        return os.name == 'nt'

    def textToSpeak(self, msg):
        pythoncom.CoInitialize()
        speak = win32com.client.Dispatch('SAPI.SpVoice')
        speak.Volume = 100
        #speak.Rate = -1
        #speak.Voice = self.speak.GetVoices('Name=Microsoft Sam').Item(0)
        speak.Speak(msg)

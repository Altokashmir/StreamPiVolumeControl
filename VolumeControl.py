"""
Per session GetMute() SetMute() GetMasterVolume() SetMasterVolume() using
SimpleAudioVolume.
"""
import getopt, sys, psutil, os
from pycaw.pycaw import AudioUtilities
import ctypes
from ctypes import wintypes



# Remove 1st argument from the
# list of command line arguments
argumentList = sys.argv[1:]
 
# Options
options = "ha:v:idm"
 
# Long options
long_options = ["Help", "Application", "Volume", "Increase", "Decrease", "Mute"]

def get_currently_focused_application():

    user32 = ctypes.windll.user32

    h_wnd = user32.GetForegroundWindow()
    pid = wintypes.DWORD()
    user32.GetWindowThreadProcessId(h_wnd, ctypes.byref(pid))
    return psutil.Process(pid.value).name()

application = get_currently_focused_application()
increase = False
decrease = False
volume_level = 0
mute = False

try:
    # Parsing argument
    arguments, values = getopt.getopt(argumentList, options, long_options)
     
    # checking each argument
    for currentArgument, currentValue in arguments:
 
        if currentArgument in ("-h", "--Help"):
            print ("usage: python VolumeControl.py [option] <value>")
            print ("   or: python VolumeControl.py [option]")
            print ("   or: python VolumeControl.py \n\n\n")

            print ("-a, --Application       Specify the name of the application that will be controled")
            print ("                        without the -a option it will use currently focused application")
            print ("-v, --Volume            Set Volume (range 0-1, scales on master volume) of application")
            print ("-i, --Increase          Increase the volume of the Application")
            print ("-d, --Decrease          Decrease the volume of the Application")
            print ("-m, --Mute              Mute the Application")
             
        elif currentArgument in ("-a", "--Application"):
            application = currentValue
            print (("Running Volume Control on % s") % (currentValue))

        elif currentArgument in ("-v", "--Volume"):
            volume_level = currentValue
            print (("Setting Application Volume to % s") % (currentValue))

        elif currentArgument in ("-i", "--Increase"):
            increase = True
            print (("Enabling special output mode % s") % (currentValue))

        elif currentArgument in ("-d", "--Decrease"):
            decrease = True
            print (("Enabling special output mode % s") % (currentValue))

        elif currentArgument in ("-m", "--Mute"):
            mute = True
             
except getopt.error as err:
    # output error, and return with an error code
    print (str(err))

class AudioController:
    def __init__(self, process_name):
        self.process_name = process_name
        self.volume = self.process_volume()

    def mute(self):
        sessions = AudioUtilities.GetAllSessions()
        for session in sessions:
            interface = session.SimpleAudioVolume
            if session.Process and session.Process.name() == self.process_name:
                interface.SetMute(1, None)
                print(self.process_name, "has been muted.")  # debug

    def unmute(self):
        sessions = AudioUtilities.GetAllSessions()
        
        for session in sessions:
            interface = session.SimpleAudioVolume
            if session.Process and session.Process.name() == self.process_name:
                interface.SetMute(0, None)
                print(self.process_name, "has been unmuted.")  # debug

    def process_volume(self):
        sessions = AudioUtilities.GetAllSessions()
        for session in sessions:
            interface = session.SimpleAudioVolume
            if session.Process and session.Process.name() == self.process_name:
                print("Volume:", interface.GetMasterVolume())  # debug
                return interface.GetMasterVolume()

    def set_volume(self, decibels):
        sessions = AudioUtilities.GetAllSessions()
        for session in sessions:
            interface = session.SimpleAudioVolume
            if session.Process and session.Process.name() == self.process_name:
                # only set volume in the range 0.0 to 1.0
                self.volume = min(1.0, max(0.0, decibels))
                interface.SetMasterVolume(self.volume, None)
                print("Volume set to", self.volume)  # debug

    def decrease_volume(self, decibels):
        sessions = AudioUtilities.GetAllSessions()
        for session in sessions:
            interface = session.SimpleAudioVolume
            if session.Process and session.Process.name() == self.process_name:
                # 0.0 is the min value, reduce by decibels
                self.volume = max(0.0, self.volume - decibels)
                interface.SetMasterVolume(self.volume, None)
                print("Volume reduced to", self.volume)  # debug

    def increase_volume(self, decibels):
        sessions = AudioUtilities.GetAllSessions()
        for session in sessions:
            interface = session.SimpleAudioVolume
            if session.Process and session.Process.name() == self.process_name:
                # 1.0 is the max value, raise by decibels
                self.volume = min(1.0, self.volume + decibels)
                interface.SetMasterVolume(self.volume, None)
                print("Volume raised to", self.volume)  # debug


def main():
    audio_controller = AudioController(application)
    if(mute):
        audio_controller.mute()
    else:
        audio_controller.unmute()

    if(increase):
        audio_controller.increase_volume(0.10) 
    
    if(decrease):
        audio_controller.decrease_volume(0.10)

if __name__ == "__main__":
    main()
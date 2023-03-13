# Import the libraries
import speech_recognition as sr #For Speech Recognition
import pygetwindow as gw #To handle the Program Windows
import subprocess as sp #To open a Program if not allready running

# Define the dictionary of programs and their paths
programs = {
    "Outlook": r"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE",
    "Word": r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE",
    "Excel": r"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE",
    "Edge": r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
}

# Listen to the microphone
r = sr.Recognizer()
mic = sr.Microphone(device_index=1)
with mic as source:
    print("Say something!")
    audio = r.listen(source)

# Analyze the audio with Google Speech Recognition
command = r.recognize_google(audio)

# Check if a program is already running. If not, start it.
running_programs = [w.title for w in gw.getAllWindows()]
program_names = programs.keys()
matched_programs = re.findall('|'.join(program_names), command)
if not matched_programs:
    print("Program not found!")
else:
    program_name = matched_programs[0]
    if program_name in running_programs:
        window = gw.getWindowsWithTitle(program_name)[0]
        window.maximize()
    else:
        program_path = programs.get(program_name)
        if program_path:
            sp.Popen([program_path])
        else:
            print("Program not found!")

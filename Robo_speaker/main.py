import win32com.client
speaker = win32com.client.Dispatch("SAPI.SpVoice")
if __name__ == '__main__':
    print("Welcome to Robo Speaker by Sultan ")
    print("Closing the program enter 'CLOSED' ")
    while True:
        s=input("Enter your words you want to say : ")
        if s == "CLOSED":
            speaker.Speak('bye bye friend')
            break
        speaker.Speak(s)

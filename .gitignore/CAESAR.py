import wx
import os
# os.environ["HTTPS_PROXY"] = "http://user:pass@192.168.1.107:3128"
import wikipedia
import wolframalpha
import time
import webbrowser
import winshell
import json
import requests
import ctypes
import random
from bs4 import BeautifulSoup
import win32com.client as wincl
from urllib.request import urlopen
import speech_recognition as sr
import ssl
# Remove SSL error
requests.packages.urllib3.disable_warnings()

try:
    _create_unverified_https_context = ssl._create_unverified_context
except AttributeError:
    # Legacy Python that doesn't verify HTTPS certificates by default
    pass
else:
    # Handle target environment that doesn't support HTTPS verification
    ssl._create_default_https_context = _create_unverified_https_context


headers = {
    'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_11_6) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/53.0.2785.143 Safari/537.36'}


speak = wincl.Dispatch("SAPI.SpVoice")

# Requirements

app_id = '8QU8RA-TE2GAVWTKL'




# GUI creation
class MyFrame(wx.Frame):
    def __init__(self):
        wx.Frame.__init__(self, None,
                          pos=wx.DefaultPosition, size=wx.Size(450, 100),
                          style=wx.MINIMIZE_BOX | wx.SYSTEM_MENU | wx.CAPTION |
                          wx.CLOSE_BOX | wx.CLIP_CHILDREN,
                          title="CAESER")
        panel = wx.Panel(self)

      

        my_sizer = wx.BoxSizer(wx.VERTICAL)
        lbl = wx.StaticText(panel,
                            label="WELOCME SIR?")
        my_sizer.Add(lbl, 0, wx.ALL, 5)
        self.txt = wx.TextCtrl(panel, style=wx.TE_PROCESS_ENTER,
                               size=(400, 30))
        self.txt.SetFocus()
        self.txt.Bind(wx.EVT_TEXT_ENTER, self.OnEnter)
        my_sizer.Add(self.txt, 0, wx.ALL, 5)
        panel.SetSizer(my_sizer)
        self.Show()
        speak.Speak("CAESAR WELCOMES YOU SIR.")
        speak.Speak("I am your personal digital assistant")
        speak.Speak("currently i cannot recognize your voice so sorry for that")
        speak.Speak("to utilize me please reed the following instructions carefully")

        print("INSTRUCTIONS")
        print("0. To view the current time just type CURRENT TIME")
        print("1. Type in your query and you will get the results")
        print("2. CAESAR cannot recognize your voice")
        print("3.Want to open a webpage type OPEN and then the webpage name for e.g. open wikipedia")
        print("4. Want to open YOUTUBE type PLAY e.g. play shape of you ")
        print("5. Want to Google anything just type search and then your query e.g search python")
        print("6.Lock your device by typing Lock the Device")
        print("7. Empty recycle bin by typing 'EMPTY' and the recycle bin will be empty")
        print("8. For entertainment purposes just type GETTING BORED and the TVF website will open for you")
        print("9.If you want to use for any other purposes you can use like weather forecasting or for calculation purposes etc.")
        

    def OnEnter(self, event):
        put = self.txt.GetValue()
        put = put.lower()
        link = put.split()
        if put == '':
            r = sr.Recognizer()
            with sr.Microphone() as src:
                audio = r.listen(src)
            try:
                put = r.recognize_google(audio)
                put = put.lower()
                link = put.split()
                self.txt.SetValue(put)

            except sr.UnknownValueError:
                print("Google Speech Recognition could not understand audio")
            except sr.RequestError as e:
                print("Could not request results from Google STT; {0}"
                      .format(e))
            except:
                print("Unknown exception occurred!")

        # Open a webpage
        if put.startswith('open '):
            try:
                speak.Speak("opening " + link[1])
                webbrowser.open('https://www.' + link[1] + '.com')
            except Exception as e:
                print(str(e))
        # Play Song on Youtube
        elif put.startswith('play '):
            try:
                link = '+'.join(link[1:])
                say = link.replace('+', ' ')
                url = 'https://www.youtube.com/results?search_query=' + link
                source_code = requests.get(url, headers=headers, timeout=15)
                plain_text = source_code.text
                soup = BeautifulSoup(plain_text, "html.parser")
                songs = soup.findAll('div', {'class': 'yt-lockup-video'})
                song = songs[0].contents[0].contents[0].contents[0]
                hit = song['href']
                speak.Speak("playing " + say)
                webbrowser.open('https://www.youtube.com' + hit)
            except Exception as e:
                print(str(e))

        # Google Search
        elif put.startswith('search '):
            try:
                link = '+'.join(link[1:])
                say = link.replace('+', ' ')
                # print(link)
                speak.Speak("searching on google for " + say)
                webbrowser.open('https://www.google.co.in/search?q=' + link)
            except Exception as e:
                print(str(e))

        # Empty Recycle bin
        elif put.startswith('empty '):
            try:
                winshell.recycle_bin().empty(confirm=False,
                                             show_progress=False, sound=True)
                print("Recycle Bin Empty!!")
            except Exception as e:
                print(str(e))
      
        # Lock the device
        elif put.startswith('lock '):
            try:
                speak.Speak("locking the device sir")
                ctypes.windll.user32.LockWorkStation()
            except Exception as e:
                print(str(e))

          # for entertainment purposes
        elif put.startswith("getting bored"):
            try:
                speak.Speak("No problem sir ")
                speak.Speak("here's something for you to enjoy") 
                webbrowser.open("https://tvfplay.com/")
            except exception as e:
                print(str(e))

        elif put.startswith("current time"):
            try:
                localtime=time.asctime(time.localtime(time.time()))
                print(localtime)
                speak.Speak(localtime)
            except Exception as e:
                print(str(e))
                
        

        # Other Cases
        else:
            # wolframalpha
            client = wolframalpha.Client(app_id)
            res = client.query(put)
            ans = next(res.results).text
            print(ans)
            speak.Speak(ans)
           


# Trigger GUI
if __name__ == "__main__":
    app = wx.App(True)
    frame = MyFrame()
    app.MainLoop()

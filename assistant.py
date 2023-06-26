from gtts import gTTS
import speech_recognition as sr
import os
import re
import webbrowser
import smtplib
import requests
from weather import Weather
import win32com.client as wincl
speak = wincl.Dispatch("SAPI.SpVoice")

def talkToMe(audio):
    #"speaks audio passed as argument"
    print(audio)
    speak.Speak(audio)
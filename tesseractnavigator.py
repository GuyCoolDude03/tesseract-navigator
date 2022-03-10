# -*- coding: utf-8 -*-
"""
Created on Mon Nov 15 17:41:18 2021

@author: guysh
"""

import numpy as np
import speech_recognition as sr
import cv2
import pytesseract
from PIL import ImageGrab
import win32api, win32con
from ctypes import windll, Structure, c_long, byref


class POINT(Structure):
    _fields_ = [("x", c_long), ("y", c_long)]


def move(x,y):
    win32api.SetCursorPos((x,y))

def click(x,y):
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,x,y,0,0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,x,y,0,0)

def get_grayscale(image):
    return cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)

def remove_noise(image):
    return cv2.medianBlur(image,5)

def thresholding(image):
    return cv2.threshold(image, 0, 255, cv2.THRESH_BINARY+ cv2.THRESH_OTSU)[1]

def upscale(image,amp):
    return cv2.resize(image, None, fx=amp, fy=amp, interpolation=cv2.INTER_CUBIC)


lookfor = "weewee"
pytesseract.pytesseract.tesseract_cmd = r'C:\Users\guysh\AppData\Local\Tesseract-OCR\tesseract.exe'

while 1:
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Say something!")
        audio = r.listen(source)
    
    try:
        # for testing purposes, we're just using the default API key
        # to use another API key, use `r.recognize_google(audio, key="GOOGLE_SPEECH_RECOGNITION_API_KEY")`
        # instead of `r.recognize_google(audio)`
        lookfor = r.recognize_google(audio).lower()
        print("Looking for " + lookfor)
    except sr.UnknownValueError:
        print("Google Speech Recognition could not understand audio")
    except sr.RequestError as e:
        print("Could not request results from Google Speech Recognition service; {0}".format(e))
    
    pt = POINT()
    windll.user32.GetCursorPos(byref(pt))
    
    x=pt.x*2
    y=pt.y*2
    
    if lookfor =="exit":
        break
    
    elif lookfor == "click":
        click(int(x),int(y))
    
    elif (lookfor == "double-click") or (lookfor == "double click"):
        click(int(x),int(y))
        click(int(x),int(y))
    
    elif lookfor == "triple click":
        click(int(x),int(y))
        click(int(x),int(y))
        click(int(x),int(y))

    else:
        filepath = 'img.png'
        
        imggrab = ImageGrab.grab()
        #imggrab = imggrab.convert('L')
        imggrab.save(filepath,'PNG')
        
        kernel = np.array([[0, -1, 0],
                       [-1, 5,-1],
                       [0, -1, 0]])
        
        img = cv2.imread('img.png')
        
        amplification = 2
        img = upscale(img,amplification)
        img = get_grayscale(img)
        img = cv2.filter2D(src=img, ddepth=-1, kernel=kernel)
        #img = cv2.GaussianBlur(img,(5,5),0)
        img = thresholding(img)
        
        pt = POINT()
        windll.user32.GetCursorPos(byref(pt))
        
        x=pt.x*2
        y=pt.y*2
        w=2
        h=2
        check = 0
        
        image_data = pytesseract.image_to_data(img,output_type=pytesseract.Output.DICT)
        
        for i, word in enumerate(image_data['text']):
            word = word.lower()
            if word != "" and (lookfor in word):
                x,y,w,h = image_data['left'][i],image_data['top'][i],image_data['width'][i],image_data['height'][i]
                #cv2.rectangle(img, (x,y), (x+w,y+h),(0,255,0),3)
                #cv2.putText(img, word,(x,y-16),cv2.FONT_HERSHEY_COMPLEX,0.25,(0,0,255), 1)
                check = 1
        
        if check == 0: print("could not be found")
        check = 0
        
        #cv2.imshow('image',img)
        move(int(x/(amplification)+w/(amplification*2)),int(y/(amplification)+h/(amplification*2)))
        #click(int(x/(amplification)+w/(amplification*2)),int(y/(amplification)+h/(amplification*2)))
        cv2.waitKey(0)


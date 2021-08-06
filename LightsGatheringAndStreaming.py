import time

import serial
#time.sleep(5.5)
import serial.tools.list_ports
import PIL
import math
from PIL import ImageGrab
from serial import serialutil
import cv2
import numpy as np
import sys

import win32gui,  win32ui,  win32con, win32api
 

#Method for capturing screen info
screenCapture = False
debugging = True

np.set_printoptions(threshold=sys.maxsize)
#Get bottom 3 rows of display
monitorRefreshRate = 165
#Num monitors you currently have connected for calculations
numMonitors = 2
#Which monitor you want to be able to write out info of
monitorToUse = 2
#The resolutions of the monitors from left to right, disregarding which one is the main monitor
monitorResolutions = np.array([[1920, 1080], [2560, 1440]])
#The positions of the monitor in accordance to where they display if they were placed on a rectangle
#For example, 2 monitors both attached at the top left, so the bottom left area is empty space, as the first monitor on the left is smaller
# monitorPositions = [[[0, 1920], [0, 1080]], [[1920,4480], [0,1440]]]
monitorPositions = np.array([[[0, 1920], [0, 1080]], [[1920,4480], [0,1440]]])
#pixelsArea = pixelArray[pixelPositioning[0,0] : pixelPositioning[1,0], pixelPositioning[0,1] : pixelPositioning[1,1]]
#determines the monitor position in the screenshot
positioning = monitorPositions[monitorToUse - 1].copy()
print(positioning)
#Horizontal offset for the pixels taken into account to calibrate closer
horizontalOffset = 0

#Number of LEDS on the light
numLEDS = 30

#Number of rows from the bottom that we want to include in our calculations, More takes a better average of the lower screen, but also takes more processing power

#Tested with a stock 5800x at 1440p, However, with a video player, the framerate drops by half if not more for some reason
#Under the windows native capture, 1 1440p row can be captured at nearly 100fps, and 5 rows at nearly 60 fps
#10 rows at 40 fps
#50 rows at about 10 fps
numRows = 1

#The exact locations were going to use for the calculations, so we grab just the pixels we want
pixelPositioning = positioning.copy()
pixelPositioning[1, 0] = pixelPositioning[1, 1] - numRows
print(pixelPositioning)


numSections = numLEDS

#Determines the resolution of the monitor being captured for internal purposes
resolutionX = (positioning[0,1] - positioning[0,0])
resolutionY = (positioning[1,1] - positioning[1,0])
print (resolutionX, resolutionY)

#Number of pixels
totalPixels = numRows * resolutionX

#Pixels per section, or light
pixelsPerSection = math.trunc(resolutionX / numSections)

#The remainder of pixels if we divide by an uneven number, this remainder will not be taken into account
pixelsRemainder = (resolutionX % numSections)

#the direction of the lightStrip
LeftToRight = True

#the arduinos port
arduinoPort = ''

#The scale of the image, if we can scale it down
scale = 1/8

# number of writes for analytics and fps
writeCount = 0

#Used for looping purposes, all start as true
notAcked = True
looping = True
notFoundAck = True

#The array that holds the total value of each pixels RGB element, is 3x the length of numPixels
avgColorArray = []

#the array that holds the average value of each pixels RGB element, is 3x the length of numPixels
pixelAvgArray = []#(numLEDS * 3)

#The array that holds the information going to be written to the lights in bytes
pixelBytes = bytearray()
#[0,0,0,0,0,0] 

#Checksum info
highBytes = ((numLEDS - 1) >> 8)
lowBytes = ((numLEDS - 1) & 0xff)

#pixelBytes += bytes(b'A') #0x41.to_bytes(1, 'big')
#pixelBytes += 0x64.to_bytes(1, 'big')
#pixelBytes += 0x61.to_bytes(1, 'big')
#pixelBytes += ((numLEDS - 1) >> 8).to_bytes(1, 'big')
#pixelBytes += ((numLEDS - 1) & 0xff).to_bytes(1, 'big')
#pixelBytes += ((highBytes ^ lowBytes ^ 0x55)).to_bytes(1, 'big')


#Acknowledgement string
ack = ''

#whether the device is connected
deviceNotConnected = True

#Find the port that the arduino is connected into, and connect to it
ports = list(serial.tools.list_ports.comports())
portTestNum = 0
for p in ports:
    if "Arduino" in p.description:
        arduinoPort = p[portTestNum]
        try:
            arduino = serial.Serial(port=arduinoPort, baudrate=115200, timeout= 10)
            print("connected to device")
            deviceNotConnected = False
            break
        except:
            print("port is currently in use, invalid, or no valid device found, retrying in 5 seconds, trying next port")
            time.sleep(5)

while(deviceNotConnected):
    try:
        arduino = serial.Serial(port=arduinoPort, baudrate=115200, timeout= 10)
        print("connected to device")
        deviceNotConnected = False
    except:
        print("port is currently in use, invalid, or no valid device found, retrying in 5 seconds")
        time.sleep(5)



#Setup our array so that they are ready
def setup():
    for i in range(0, (numLEDS) * 3 + 3,  1):
        avgColorArray.append(0)
        pixelAvgArray.append(0)
        #pixelBytes.append(0)
    print("completed setup")    


#The main loop sequence to gather the pixel information, calculate averages, and write those averages to bytes
def loop():
    #Checksum
    pixelBytes = bytearray()
    pixelBytes += bytes(b'A') #0x41.to_bytes(1, 'big')
    pixelBytes += 0x64.to_bytes(1, 'big')
    pixelBytes += 0x61.to_bytes(1, 'big')
    pixelBytes += ((numLEDS - 1) >> 8).to_bytes(1, 'big')
    pixelBytes += ((numLEDS - 1) & 0xff).to_bytes(1, 'big')
    pixelBytes += ((highBytes ^ lowBytes ^ 0x55)).to_bytes(1, 'big')

    #refresh the total pixel color array to start from 0
    for i in range(0, (numLEDS) * 3 + 3,  1):
        avgColorArray[i] = 0

    #for calculating amount of time it takes to gather the screenshot
    #t0 = time.process_time()
    #print(t0)

    #A method is chosen depending on the options listed above, there are limitations to each method.
    #Each method fills the pixelArray with our pixel information for the entire screen, so we can grab that information later for processing

    #Slow method, takes roughly .03 seconds to gather a frame of information for the desktop at 2560x1440, on a Ryzen 5800x, resulting in about 20-30fps MAX
    #This method works on Windows and Mac, but not Linux
    if screenCapture:
        image = PIL.ImageGrab.grab()
        #image.show()
        #image = image.resize((math.trunc(resolutionX * scale), math.trunc(resolutionY * scale)), PIL.Image.BICUBIC)
        #for y in range(0, numRows, 1):
            #for x in range(0, math.trunc(resolutionX * scale), 1):
                #color = image.getpixel((x, math.trunc((resolutionY - 1)) * scale - y))
        pixelArray = np.asarray(image)
        #print(pixelArray.shape)
        pixelsArea = pixelArray[resolutionY - numRows : resolutionY, :, :]
        #print(pixelsArea.shape)
        #print(pixelsArea)
        pixelsArea = pixelsArea.flatten()
        pixelsArea = pixelsArea.reshape((-1,3))
        #print(pixelsArea)
                    #if (x % pixelsPerSection != 0):
                        #pixelArray.append(color)

    #This method Is supposedly magnitudes faster, allowing for almost instant capture at high frame rates, however is only available on Windows
    elif(debugging):
        hDesk = win32gui.GetDesktopWindow()
        #(hDesk)
        
        # you can use this to capture only a specific window
        #l, t, r, b = win32gui.GetWindowRect(hwnd)
        #w = r - l
        #h = b - t
        
        # get complete virtual screen including all monitors
        SM_XVIRTUALSCREEN = 76
        SM_YVIRTUALSCREEN = 77
        SM_CXVIRTUALSCREEN = 78
        SM_CYVIRTUALSCREEN = 79
        w = vscreenwidth = win32api.GetSystemMetrics(SM_CXVIRTUALSCREEN)
        h = vscreenheigth = win32api.GetSystemMetrics(SM_CYVIRTUALSCREEN)
        l = vscreenx = win32api.GetSystemMetrics(SM_XVIRTUALSCREEN)
        t = vscreeny = win32api.GetSystemMetrics(SM_YVIRTUALSCREEN)
        r = l + w
        b = t + h
        
        #print (l, t, r, b, ' -> ', w, h)
        
        hDeskDC = win32gui.GetWindowDC(hDesk)
        imgDC  = win32ui.CreateDCFromHandle(hDeskDC)
        memDC = imgDC.CreateCompatibleDC()
        
        saveBitMap = win32ui.CreateBitmap()
        saveBitMap.CreateCompatibleBitmap(imgDC, math.trunc(r * scale), numRows)
        memDC.SelectObject(saveBitMap)

        #memDC.BitBlt((0,0),
         #   (r, b - (b - numRows)),
          #  imgDC,
           # (0, b - numRows),
            #win32con.SRCCOPY)

        memDC.StretchBlt((0, 0),
            (math.trunc(r * scale), math.trunc((b  - (b - (math.floor(numRows)))))),
            imgDC,
            (0, b - math.floor(numRows / scale)),
            (round(w * .5703125), b - (b - math.trunc(numRows / scale))),
            win32con.SRCCOPY)

        signedIntsArray = saveBitMap.GetBitmapBits(True)
        img = np.fromstring(signedIntsArray, dtype='uint8')
        img.shape = (math.trunc(r * scale), numRows, 4)
        #saveBitMap.SaveBitmapFile(memDC,  'screencapture.bmp')
        #time.sleep(3)

        # Free Resources
        #hDeskDC.DeleteDC()
        imgDC.DeleteDC()
        memDC.DeleteDC()
        win32gui.ReleaseDC(hDesk, hDeskDC)
        win32gui.DeleteObject(saveBitMap.GetHandle())

                # drop the alpha channel, or cv.matchTemplate() will throw an error like:
                #   error: (-215:Assertion failed) (depth == CV_8U || depth == CV_32F) && type == _templ.type() 
                #   && _img.dims() <= 2 in function 'cv::matchTemplate'
        img = img[...,:3]
        imgTest = np.flip(img, axis=2)
        #print(imgTest)
        #img2 = img[...,:0]
        #img3 = img[...,1:1]
        #img4 = img[...,2:2]
        #imgFinal = np.hstack((img4, img3, img2))
        #img = img[:, resolutionY - (numRows + 1) : resolutionY - 1]


        # make image C_CONTIGUOUS to avoid errors that look like:
                #   File ... in draw_rectangles
                #   TypeError: an integer is required (got type tuple)
                # see the discussion here:
                # https://github.com/opencv/opencv/issues/14866#issuecomment-580207109
        #imgFinal = np.ascontiguousarray(imgFinal)
        #print(imgFinal)
        pixelArray = imgTest
        pixelsArea = pixelArray.flatten().reshape(-1,3)

        #print(pixelsArea)
        #time.sleep(1)
        #print (pixelsArea)
        
    #Capture the screen via video device. If you have a capture card, this would be beneficial, as it is supposedly also faster.
    #Available on whatever device can access your video feed, but difficult to setup on Raspberry Pi
    else:
        vid = cv2.VideoCapture(0)
        image = vid.read()
        #for y in range(0, numRows, 1):
            #for x in range(0, math.trunc(resolutionX * scale), 1):
                #color = image[x, math.trunc((resolutionY - 1) * scale) - y]
        pixelArray = np.asarray(image)

    #t0 = time.process_time()
    #print(t0)
    #image.show() 
    #print(pixelArray)
    #print(pixelArray.shape)
    #print( pixelArray[pixelPositioning[0,0] - 1 : pixelPositioning[1,0] - 1, pixelPositioning[0,1] - 1 : pixelPositioning[1,1] - 1,:])
    #pixelsArea = pixelArray[pixelPositioning[1,0] - 1: pixelPositioning[1,1] - 1, pixelPositioning[0,0] - 1: pixelPositioning[0,1] - 1, :]
    #pixelsArea = pixelsArea.flatten().flatten()
    #print(pixelsArea.shape)
    #The index of the color of the pixel we are on
    #print(pixelsArea)
    pixelColor = horizontalOffset
    #numPixel = abs(horizontalOffset)
    #TODO pixelOffset handling
    #print("Screen Grabbed")
    #print(pixelsArea)
    #pixelsAreaChecker = []
    #add each pixel being looked at into the avgColorArray for the pixel it will light up
    for pixel in pixelsArea:
        #pixelsAreaChecker.append(pixel)
        if pixelColor < 0: #or color > resolutionX:
            pixelColor += 1
            #numPixel += 1
        #elif debugging:
            #avgColorArray[math.trunc(pixelColor / (pixelsPerSection * numRows * 3)) % (numSections * 3)] = avgColorArray[math.trunc(pixelColor / (pixelsPerSection * numRows * 3)) % (numSections * 3)] + pixel
            #pixelColor += 1
        #else: 
            #avgColorArray[math.trunc(pixelColor / (pixelsPerSection * 3)) % (numSections * 3)] = avgColorArray[math.trunc(pixelColor / (pixelsPerSection * 3)) % (numSections * 3)] + pixel
            #pixelColor += 1
            #avgColorArray[(math.trunc(pixelColor / math.trunc(pixelsPerSection * scale)) % numSections) * 3] = avgColorArray[(math.trunc(pixelColor / math.trunc(pixelsPerSection * scale)) % numSections) * 3] + pixel[0]
        else:
            index = (math.trunc(pixelColor / math.trunc(pixelsPerSection * scale)) % numSections)
            avgColorArray[index * 3] = avgColorArray[index * 3] + pixel[0]
            avgColorArray[index * 3 + 1] = avgColorArray[index * 3 + 1] + pixel[1]
            avgColorArray[index * 3 + 2] = avgColorArray[index * 3 + 2] + pixel[2]
            pixelColor += 1
            #numPixel += 1
    #print(avgColorArray)
    #print(pixelsAreaChecker)

    #For each of the average colors, divide it by the number of pixels in the section, move it into the pixelAvgArray, and then convert it to bytes in our pixelBytes array
    for avgColor in range((max(0, horizontalOffset)) + 6, (len(avgColorArray) - abs(horizontalOffset)) + 6, 1):
        pixelAvgArray[avgColor - 6] = (math.trunc(avgColorArray[avgColor - 6] / (pixelsPerSection * numRows)))
        #print(pixelAvgArray)
        pixelBytes += (pixelAvgArray[avgColor - 6].to_bytes(1, 'big'))
        #print(str(pixelAvgArray[avgColor - 6]) + ' at ' + str(avgColor - 6))

    #write out debugging information every 10 frames processed, as its easier to look at if you have little changing on the screen
    if (writeCount % 10 == 0):
        #print(avgColorArray)
        #print(pixelAvgArray)
        #print(len(pixelAvgArray))
        print(writeCount)


    #print(pixelBytes)
    #data = bytes(pixelBytes)
    #print(to_Bytes(pixelBytes))

    #arr = []
    #for listNum in range(255):
        #arr.append(listNum.to_bytes(1, 'big'))
    #print(arr)
    #print("Filled array")
    #print(time.process_time())
    #time.sleep(5.5)

    #Finally, write the information to the device
    write_read(pixelBytes)


#Write the information to the device, and make sure that the device acknowledges this script, and is setup correctly
def write_read(data):
    #print("entered Write")
    if notAcked:
        #print("notAcked is true")
        notFoundAck = True
        while notFoundAck:
            #print("notFoundAck is true")
            #Read our arduinos data to make sure there is acknoledgement
            ack = arduino.readline()
            #print("ack is " + ack.decode('utf-8'))
            #if we have acknowledgement write, and save that weve been acknowledged
            if "Ada" in ack.decode('utf-8'):
                notFoundAck = False
                arduino.write(data)
    #Once we have been acknowledged, write to the device
    else:
        #print("notAcked is false, immediate writing") 
        arduino.write(data)
    #print("Finished Write")

    #time.sleep(1/monitorRefreshRate)

#def resetData():
    #pixelBytes = []
    #pixelAvgArray = []
    #avgColorArray = []


#The main process of the script
#Run setup and then do the loop infinitely
setup()

while(looping):
    loop()
    #time.sleep(10.8)
    writeCount += 1
    #print(writeCount)
    notAcked = False

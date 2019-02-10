import socket
import sys
from pptController import pptController
HOST = ''  # Symbolic name, meaning all available interfaces
PORT = 8888  # Arbitrary non-privileged port

s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
print 'Socket created'

# Bind socket to local host and port
try:
    s.bind((HOST, PORT))
except socket.error as msg:
    print 'Bind failed. Error Code : ' + str(msg[0]) + ' Message ' + msg[1]
    sys.exit()

print 'Socket bind complete'

# Start listening on socket
s.listen(10)
print 'Socket now listening'

# now keep talking with the client
while 1:

    conn, addr = s.accept()
    print 'Connected with ' + addr[0] + ':' + str(addr[1])
    msg = conn.recv(1024)
    print (msg)
    if msg:
        ppt=pptController()
        split=","
        if msg.endswith("Start"):   #start the presentation
            ppt.fullScreen()
           # r="thankyou"
            #conn.send(r)
            m=ppt.getActivePresentationSlideIndex()
            n=ppt.getActivePresentationSlideCount()
            conn.send(str(m)+split+str(n))
        if msg.endswith("Pause"):       # pause the presentation
            ppt.click()
            m = ppt.getActivePresentationSlideIndex()
            n = ppt.getActivePresentationSlideCount()
            conn.send(str(m) +split+ str(n))

        if msg.endswith("Next"):    # go to next slide
            ppt.nextPage()
            m = ppt.getActivePresentationSlideIndex()
            n = ppt.getActivePresentationSlideCount()
            conn.send(str(m) +split+ str(n))
        if msg.endswith("Previous"):    # go to previous slide
            ppt.prePage()
            m = ppt.getActivePresentationSlideIndex()
            n = ppt.getActivePresentationSlideCount()
            conn.send(str(m) +split+ str(n))

        if msg[-1].isdigit():    # go to specific slide
            if msg[-2].isdigit():
                c=int(msg[-2:])
            else:
                c=int(msg[-1])
            print c
            ppt.gotoSlide(c)
            m = ppt.getActivePresentationSlideIndex()
            n = ppt.getActivePresentationSlideCount()
            conn.send(str(m) +split+ str(n))

s.close()

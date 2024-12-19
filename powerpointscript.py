#Python3

import win32com.client

def main():
    ppt = win32com.client.Dispatch("PowerPoint.Application")
    ppt.Activate()
    ppt.Visible = True
    present = ppt.Presentations.Open("C:\\Users\\Obiwa\\Desktop\\defaultpresentation.pptx")
    
    # present.SlideShowSettings.Run()
if "__main__" == __name__:
    main()

#Python3

import win32com.client

def main():
    ppt = win32com.client.Dispatch("PowerPoint.Application")
    ppt.Activate()
    ppt.Visible = True
    present = ppt.Presentations.Open("C:\\Users\\Obiwa\\Documents\\defaultpresentation.pptx.pptx")
    
    
    pptsettings = present.SlideShowSettings
    pptsettings.ShowPresenterView = False

    pptsettings.Run()
if "__main__" == __name__:
    main()

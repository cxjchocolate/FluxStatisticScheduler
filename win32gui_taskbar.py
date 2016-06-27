# Creates a task-bar icon.  Run from Python.exe to see the
# messages printed.
import win32api, win32gui
import win32con, winerror
import sys, os, subprocess
import env
import transfer
import logging

class MainWindow:
    def __init__(self):
        msg_TaskbarRestart = win32gui.RegisterWindowMessage("TaskbarCreated");
        message_map = {
                msg_TaskbarRestart: self.OnRestart,
                win32con.WM_DESTROY: self.OnDestroy,
                win32con.WM_COMMAND: self.OnCommand,
                win32con.WM_USER+20 : self.OnTaskbarNotify,
        }
        # Register the Window class.
        wc = win32gui.WNDCLASS()
        hinst = wc.hInstance = win32api.GetModuleHandle(None)
        wc.lpszClassName = "SchedulerTaskbar"
        wc.style = win32con.CS_VREDRAW | win32con.CS_HREDRAW;
        wc.hCursor = win32api.LoadCursor( 0, win32con.IDC_ARROW )
        wc.hbrBackground = win32con.COLOR_WINDOW
        wc.lpfnWndProc = message_map # could also specify a wndproc.

        # Don't blow up if class already registered to make testing easier
        try:
            classAtom = win32gui.RegisterClass(wc)
        except win32gui.error as err_info:
            if err_info.winerror!=winerror.ERROR_CLASS_ALREADY_EXISTS:
                raise

        # Create the Window.
        style = win32con.WS_OVERLAPPED | win32con.WS_SYSMENU
        self.hwnd = win32gui.CreateWindow( wc.lpszClassName, "Scheduler Taskbar", style, \
                0, 0, win32con.CW_USEDEFAULT, win32con.CW_USEDEFAULT, \
                0, 0, hinst, None)
        win32gui.UpdateWindow(self.hwnd)
        self._DoCreateIcons()
        self._DoCreate()
        
    def _DoCreateIcons(self):
        # Try and find a custom icon
        hinst =  win32api.GetModuleHandle(None)
        
        path = env.getHome()
        
        iconPathName = os.path.abspath(os.path.join( path, "sched.ico" ))
        if os.path.isfile(iconPathName):
            icon_flags = win32con.LR_LOADFROMFILE | win32con.LR_DEFAULTSIZE
            hicon = win32gui.LoadImage(hinst, iconPathName, win32con.IMAGE_ICON, 0, 0, icon_flags)
        else:
            print("Can't find a Python icon file - using default")
            hicon = win32gui.LoadIcon(0, win32con.IDI_APPLICATION)

        flags = win32gui.NIF_ICON | win32gui.NIF_MESSAGE | win32gui.NIF_TIP
        nid = (self.hwnd, 0, flags, win32con.WM_USER+20, hicon, "Scheduler")
        try:
            win32gui.Shell_NotifyIcon(win32gui.NIM_ADD, nid)
        except win32gui.error:
            # This is common when windows is starting, and this code is hit
            # before the taskbar has been created.
            print("Failed to add the taskbar icon - is explorer running?")
            # but keep running anyway - when explorer starts, we get the
            # TaskbarCreated message.

    def OnRestart(self, hwnd, msg, wparam, lparam):
        self._DoCreateIcons()
        
    def _DoCreate(self):
        #transfer function
        transfer.main(False)

    def OnDestroy(self, hwnd, msg, wparam, lparam):
        nid = (self.hwnd, 0)
        win32gui.Shell_NotifyIcon(win32gui.NIM_DELETE, nid)
        win32gui.PostQuitMessage(0) # Terminate the app.

    def OnTaskbarNotify(self, hwnd, msg, wparam, lparam):
        if lparam==win32con.WM_RBUTTONUP:
            menu = win32gui.CreatePopupMenu()
            win32gui.AppendMenu( menu, win32con.MF_STRING, 1024, "Show Log")
            win32gui.AppendMenu( menu, win32con.MF_STRING, 1026, "About Us")
            win32gui.AppendMenu( menu, win32con.MF_STRING, 1027, "Option")
            win32gui.AppendMenu( menu, win32con.MF_STRING, 1028, "Upload Offline")
            win32gui.AppendMenu( menu, win32con.MF_STRING, 1025, "Exit program" )
            pos = win32gui.GetCursorPos()
            # See http://msdn.microsoft.com/library/default.asp?url=/library/en-us/winui/menus_0hdi.asp
            win32gui.SetForegroundWindow(self.hwnd)
            win32gui.TrackPopupMenu(menu, win32con.TPM_LEFTALIGN, pos[0], pos[1], 0, self.hwnd, None)
            win32gui.PostMessage(self.hwnd, win32con.WM_NULL, 0, 0)
        return 1

    def OnCommand(self, hwnd, msg, wparam, lparam):
        id = win32api.LOWORD(wparam)
        if id == 1024:
            subprocess.call (["notepad", "log.log"])
        elif id == 1025:
            win32gui.DestroyWindow(self.hwnd)
        elif id == 1026:
            subprocess.call (["notepad", "readme.txt"])
        elif id == 1027:
            subprocess.call (["notepad", "conf/conf.ini"])
        elif id == 1028:
            transfer.transferHistDetails(env.config.get("DB_FILE"), env.config.get("STOREID"), 
                         env.config.get("FTP_HOME"), env.config.get("FTP_HOST"),
                         env.config.get("FTP_PORT"), env.config.get("FTP_USERNAME"),
                         env.config.get("FTP_PASSWORD"))
        else:
            print("Unknown command -", id)

def main():
    w=MainWindow()
    win32gui.PumpMessages()
    

if __name__=='__main__':
    logging.basicConfig(level=logging.DEBUG,
                format='%(asctime)s [%(levelname)s] %(filename)s[line:%(lineno)d] %(message)s',
                datefmt='%a, %d %b %Y %H:%M:%S', filename="log.log")
    console = logging.StreamHandler()
    console.setLevel(logging.DEBUG)
    formatter = logging.Formatter('%(name)-12s: %(levelname)-8s %(message)s')
    console.setFormatter(formatter)
    logging.getLogger('').addHandler(console)
    main()
    

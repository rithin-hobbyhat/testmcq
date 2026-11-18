"""
Create Desktop and folder shortcuts with the custom 'R' icon.
"""
import os
import sys
import subprocess

def create_windows_shortcut():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    ico_path = os.path.join(base_dir, "app_icon.ico")
    app_script = os.path.join(base_dir, "app.py")
    
    # Path to user desktop
    desktop_dir = os.path.expanduser("~/Desktop")
    if not os.path.exists(desktop_dir):
        # Fallback to standard Desktop path
        desktop_dir = r"C:\Users\Windows 10 Pro\Desktop"

    # Shortcuts to create
    desktop_shortcut = os.path.join(desktop_dir, "MOQ Question Bank.lnk")
    folder_shortcut = os.path.join(base_dir, "MOQ Question Bank.lnk")

    # Find pyw.exe or pythonw.exe for silent launch (no black terminal window)
    pythonw_exe = sys.executable.replace("python.exe", "pythonw.exe")
    if not os.path.exists(pythonw_exe):
        pythonw_exe = "pythonw.exe"

    # PowerShell script to create shortcuts via WScript.Shell
    ps_script = f"""
    $WshShell = New-Object -comObject WScript.Shell
    
    # 1. Create Shortcut on Desktop
    $Shortcut = $WshShell.CreateShortcut("{desktop_shortcut}")
    $Shortcut.TargetPath = "pyw.exe"
    $Shortcut.Arguments = '"{app_script}"'
    $Shortcut.WorkingDirectory = '"{base_dir}"'
    $Shortcut.IconLocation = '"{ico_path}"'
    $Shortcut.Description = "MOQ Question Bank Authoring Studio"
    $Shortcut.Save()

    # 2. Create Shortcut in MOQ Folder
    $ShortcutFolder = $WshShell.CreateShortcut("{folder_shortcut}")
    $ShortcutFolder.TargetPath = "pyw.exe"
    $ShortcutFolder.Arguments = '"{app_script}"'
    $ShortcutFolder.WorkingDirectory = '"{base_dir}"'
    $ShortcutFolder.IconLocation = '"{ico_path}"'
    $ShortcutFolder.Description = "MOQ Question Bank Authoring Studio"
    $ShortcutFolder.Save()
    """

    cmd = ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps_script]
    res = subprocess.run(cmd, capture_output=True, text=True)
    if res.returncode == 0:
        print(f"Created Desktop Shortcut: {desktop_shortcut}")
        print(f"Created Folder Shortcut: {folder_shortcut}")
    else:
        print(f"Error creating shortcut: {res.stderr}")

    # Also create a VBS silent launcher
    vbs_path = os.path.join(base_dir, "Launch_Silent.vbs")
    with open(vbs_path, "w", encoding="utf-8") as f:
        f.write(f'CreateObject("Wscript.Shell").Run "pyw.exe ""{app_script}""", 0, False\n')
    print(f"Created silent VBS launcher: {vbs_path}")

if __name__ == "__main__":
    create_windows_shortcut()

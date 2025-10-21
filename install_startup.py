"""
Creates a shortcut in the current user's Windows Startup folder to run the print agent on login.

Run on Windows as the user who should auto-run the agent.
"""
import os
import sys
from pathlib import Path

def create_shortcut(target_py, args='', name='print_agent'):
    try:
        # pywin32 provides the COM wrappers used to create a .lnk shortcut
        import pythoncom  # noqa: F401
        from win32com.shell import shell
        from win32com.client import Dispatch
    except Exception:
        print('pywin32 is required to create a Windows shortcut. Install with: pip install pywin32')
        raise

    # Attempt to get the user's Startup folder; fall back to APPDATA path
    try:
        # Prefer the shell API when available
        startup = shell.SHGetFolderPath(0, 7, None, 0)  # CSIDL_STARTUP == 7
    except Exception:
        startup = os.path.join(os.environ.get('APPDATA', ''), 'Microsoft', 'Windows', 'Start Menu', 'Programs', 'Startup')

    if not os.path.isdir(startup):
        raise RuntimeError('Startup folder not found: ' + str(startup))

    shortcut_path = os.path.join(startup, f"{name}.lnk")

    shell_link = Dispatch('WScript.Shell').CreateShortcut(shortcut_path)
    shell_link.TargetPath = sys.executable
    shell_link.Arguments = f'"{target_py}" {args}'
    shell_link.WorkingDirectory = str(Path(target_py).parent)
    shell_link.IconLocation = sys.executable
    shell_link.save()
    return shortcut_path


if __name__ == '__main__':
    if os.name != 'nt':
        print('This script is for Windows only')
        sys.exit(1)

    if len(sys.argv) >= 2:
        target = sys.argv[1]
    else:
        # assume app.py in the same folder
        target = os.path.join(os.path.dirname(__file__), 'app.py')

    args = ' '.join(sys.argv[2:])

    try:
        path = create_shortcut(target, args=args)
        print('Created shortcut:', path)
    except Exception as e:
        print('Failed to create shortcut:', e)

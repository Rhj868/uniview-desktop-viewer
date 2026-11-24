"""UniView silent launcher - double-click to start with no console window.

Windows opens .pyw files with pyw.exe (no console). server.py already
redirects stdout/stderr to uniview.log when there is no console, and opens
the UniView window itself. No extra process, no black box.
"""
import os
import runpy

_ROOT = os.path.dirname(os.path.abspath(__file__))
runpy.run_path(os.path.join(_ROOT, "server.py"), run_name="__main__")

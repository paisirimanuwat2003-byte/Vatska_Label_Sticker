# utils.py
import os
import sys

def resource_path(relative_path):
    """ Get the absolute path to a resource, works for dev and for PyInstaller """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def get_save_path(filename, foldername):
    """Creates the Extracted_Data folder if needed and returns the full path string."""
    os.makedirs(foldername, exist_ok=True)
    return os.path.join(foldername, filename)
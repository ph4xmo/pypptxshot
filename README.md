# 📷 PyPPTXShot - A Python Screenshotting PPTX Generator
This small Python app serves as an automatic screenshotter (and 
optional auto-clicker) of a selected area in your screen with the help 
of PyAutoGUI and python-pptx. This is also uses Tkinter as its GUI.

PyPPTXShot was initially a past-time project to capture 
slideshow presentations that are undownloadable (probably because of a 
certain professor won't let it be downloadable for some reason during 
online class) and is displayed in a website instead (e.g., ISpring 
Presentation on Moodle).

The default save path of the generated PPTX file would be in the
desktop. While the default (and initial) filename is "pypptxshot_output-0". 
Best used with a console to view what is going on.

Contributions are welcomed!

## ✨ Features
- Optional auto-clicker (to simulate "next" on target slides).
- Auto-clicker can also have a click delay.
- User can have any number of screenshots per slide (tested up to 500).
- Placed screenshots can be scaled.
- Maximum slide size is limited by python-pptx.
- Slide size can be automatically adjusted based from screenshot size.
- F2 shortcut key for getting coordinates in screen (for auto-click).

## 📦 Installation
```
pip install pypptxshot
python -m pypptxshot
```

## 🔑 License
PyPPTXShot is released under the MIT License. See LICENSE for more details."# pypptxshot" 

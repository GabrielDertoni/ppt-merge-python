# Merge PowerPoint Presentations

## Description
A program that merges two or more PowerPoint presentations into one.

### Setup
1. Download or clone this git repo.
2. Download the tkdnd binaries from [Source Forge (32 bits)](http://sourceforge.net/projects/tkdnd) or [Totally Random Website that I've Found (64 bits)](https://osdn.net/projects/sfnet_tkdnd/downloads/Windows%20Binaries/TkDND%202.8/tkdnd2.8-win32-x86_64.tar.gz/) and put it inside the `tcl/` folder in your python instalation. In my case, I had a virtualenv and heance, I moved the `tkdnd2.8/` folder inside `env/tcl/`.
3. Download the TkinterDnD wrapper class at https://sourceforge.net/projects/tkinterdnd/ and put the `TkinterDnD2/` folder inside your Python's `Lib/site-packages/` folder.
4. Install all required dependencies in the `requirements.txt` file.

### Getting started
Run `python main.py <presentation path> --name <output path>`, replacing the "\<presentation path\>" with two or more paths to PowerPoint presentations (.pptx files). \<output path\> should be substituted with the path to save the newly created presentation to.

If you run `python app.py` a GUI will appear in which you can use the same functionality as before, exept in a graphical way. It should be preatty simple and intuitive.

### Files and their functions
The `main.py` file is a working program on its own. It can merge multiple PowerPoint presentations on the command line.
The `app.py` file is a program that utilizes the `main.py` tool with a Grafical User Interace (GUI), mainly implemented in `gui.py`. `presentation_file.py` is just a helper file.

## Compiling to .exe
To the code and create an executable, use `pyinstaller -i icon.ico -w --add-data="icon.ico;." --add-data="<path to tkdnd2.8>;tcl/tkdnd2.8" app.py`
If there is a virtualenv in the project folder called __ppt-merge__ and tkdnd2.8 is inside of its `tcl/` folder, the same compiling can be done with only `pyinstaller app.spec` and everything should work fine.
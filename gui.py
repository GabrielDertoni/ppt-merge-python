import os
import re
from tkinter import *
from tkinter import filedialog
from TkinterDnD2 import *
from threading import Thread

from main import merge_presentations
from presentation_file import PresentationFile


def merge(output_entry, message_label):
  output_path = output_entry.get()
  vars_check = PresentationFile.check()
  output_check = os.path.exists(os.path.split(output_path)[0]) and not os.path.exists(output_path) and os.path.splitext(output_path)[1] == ".pptx"
  if vars_check and output_check:
    presentations = PresentationFile.get_filepaths()
    def wrapper_worker(pres, out):
      merge_presentations(pres, out)
      message_label.config(text="Merged presentations successfully!", fg="green")
    
    thread = Thread(target=wrapper_worker, args=(presentations, output_path))
    thread.start()
    message_label.config(text="Merging presentations...", fg="blue")
  elif not vars_check:
    message_label.config(text="An error occurred in the presentation input fields...", fg="red")
  else:
    message_label.config(text="An error occurred in the output field...", fg="red")

def add_input(frame, deletable=True):
  PresentationFile(frame, deletable=deletable)

def setup_DnD(frame, input_frame):
  # Make the window a drop target for files.
  frame.drop_target_register(DND_FILES)
  def drop(event):
    # Detect if one or several files are beeing droped
    # print(event.data)
    # Match any number of characters between brackets.
    # Any character: [\s\S]; One or more characters: *; Match as few characters as possible: ?; Group: (); In brackets: {}
    regExp = r"{([\s\S]*?)}"
    with_whitespace = re.findall(regExp, event.data)
    no_whitespace = [path for path in re.sub(regExp, "", event.data).split(" ") if path != ""]
    filepaths = with_whitespace + no_whitespace
    # Get the last empty input, from bottom to top. Those empty imputs will be filled with the dragged files.
    last_empty_var = len(PresentationFile.inputs)
    for _ in range(last_empty_var):
      if last_empty_var >= 1:
        if PresentationFile.get_value(last_empty_var-1) == "":
          last_empty_var -= 1
          continue
      break

    for i, fpath in enumerate(filepaths):
      if i >= last_empty_var and i < len(PresentationFile.inputs):
        PresentationFile.set_value(i, fpath)
      else:
        add_input(input_frame)
        PresentationFile.set_value(-1, fpath)
  
  frame.dnd_bind('<<Drop>>', drop)

def render_gui():
  # root = Tk()
  root = TkinterDnD.Tk()
  root.title("PowerPoint Merger")
  root.geometry("700x700")
  root.iconbitmap(os.path.abspath("icon.ico"))

  ## Top heading for the hole view.
  heading = Label(text="PowerPoint Merger", bg="#d04424", fg="white", height="2", font=("TkHeadingFont", 30))
  heading.pack(side=TOP, fill=X)

  ## Configure the output path input frame, label and entry ##
  output_frame = Frame(root, pady="30")
  output_frame.pack(side=TOP)
  # Label
  output_label_frame = Frame(output_frame)
  output_label_frame.pack(side=TOP, fill=X)
  output_label = Label(output_label_frame, text="Output file path", font="Helvetica 10 bold")
  output_label.pack(side=LEFT)
  # Entry
  output_entry_frame = Frame(output_frame)
  output_entry_frame.pack(side=TOP, fill=X)
  output_path = StringVar()
  output_entry = Entry(output_entry_frame, textvariable=output_path, width="30")
  output_entry.pack(side=LEFT)
  # Browse button
  def save_dialog():
    filename = filedialog.asksaveasfilename(initialdir="/", title="Save as...", filetype=[("PowerPoint files", "*.pptx")])
    if os.path.splitext(filename)[1] == '': filename += '.pptx'
    output_path.set(filename)
    output_entry.xview_moveto(1)

  output_browse = Button(output_entry_frame, text="Browse", command=save_dialog)
  output_browse.pack(side=RIGHT)

  ## Label for input files ##
  input_label = Label(root, text="Input PowerPoint presentations", bg="#d04424", fg="white", font="Helvetica 12 bold")
  input_label.pack(side=TOP, fill=X)

  ## Create frame to store all presentation inputs
  canvas_frame = Frame(root, pady="15")
  canvas_frame.pack(side=TOP)

  # Make the input frame scrollable
  canvas = Canvas(canvas_frame, highlightthickness=0)
  input_frame = Frame(canvas)
  scrollbar = Scrollbar(canvas_frame, orient="vertical", command=canvas.yview)
  canvas.configure(yscrollcommand=scrollbar.set)

  def update_canvas(event):
    canvas.configure(scrollregion=canvas.bbox("all"), width=400, height=300)

  scrollbar.pack(side=RIGHT, fill=Y)
  canvas.pack(side=LEFT)
  canvas.create_window((210, 0), window=input_frame)
  input_frame.bind("<Configure>", update_canvas)

  # Put the two input fields for merging the presentations (at least two)
  add_input(input_frame, deletable=False)
  add_input(input_frame, deletable=False)

  add_btn_frame = Frame(root)
  add_btn_frame.pack(side=TOP)

  setup_DnD(root, input_frame)

  # Button to add a new presentation to merge
  add_btn = Button(add_btn_frame, text="Add presentation", command=lambda: add_input(input_frame), bg="#d04424", fg="white")
  add_btn.pack(side=LEFT, pady="10", padx="5")

  def browse_many_dialog():
    filenames = filedialog.askopenfilenames(title="Select presentations", filetype=[("PowerPoint files", "*.pptx")])
    if len(PresentationFile.inputs) == 2 and PresentationFile.get_value(0) == "" and PresentationFile.get_value(1) == "" and len(filenames) >= 2:
      for i in range(len(filenames)):
        if i < 2:
          PresentationFile.set_value(i, filenames[i])
        else:
          add_input(input_frame)
          PresentationFile.set_value(-1, filenames[i])
    else:
      for fname in filenames:
        add_input(input_frame)
        PresentationFile.set_value(-1, fname)

  browse_btn = Button(add_btn_frame, text="Select files", command=browse_many_dialog, bg="#d04424", fg="white")
  browse_btn.pack(side=RIGHT, pady="10", padx="5")

  # Label where all errors are displayed
  error_frame = Frame(root)
  error_frame.pack(side=TOP)
  error_label = Label(error_frame, text="", fg="red", font="Consolas 8 bold")
  error_label.pack(side=TOP)

  # Button to merge the presentations
  button = Button(root, text="Merge", command=lambda: merge(output_path, error_label), width="30", bg="#1c5aab", fg="white", font="Helvetica 12 bold")
  button.pack(side=BOTTOM, pady="15")

  root.mainloop()
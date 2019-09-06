from tkinter import *
from tkinter import filedialog
import os
from threading import Thread

from main import merge_presentations

root = Tk()
root.title("PowerPoint Merger")
root.geometry("700x700")
root.iconbitmap(os.path.abspath("icon.ico"))

vars = []
presentation_entries = []

def input_layout(frame, var, deletable=True):
  vars_len = len(vars)
  background = "lightgrey" if (vars_len + 1) % 2 == 0 else "SystemButtonFace"
  container_frame = Frame(frame, bg=background, pady="5", padx="5")
  container_frame.pack(side=TOP, fill=X, pady="5", padx="5")
  label_frame = Frame(container_frame, bg=background)
  label_frame.pack(side=TOP, fill=X)
  entry_frame = Frame(container_frame, bg=background)
  entry_frame.pack(side=TOP, fill=X)
  arrows_frame = Frame(entry_frame, bg=background)
  arrows_frame.pack(side=LEFT, fill=Y, padx="5")

  def swap_vars(direction):
    if direction == "up" and vars_len - 1 >= 0:
      buff = var.get()
      var.set(vars[vars_len - 1].get())
      vars[vars_len - 1].set(buff)
    elif direction == "down" and vars_len + 1 < len(vars):
      buff = var.get()
      var.set(vars[vars_len + 1].get())
      vars[vars_len + 1].set(buff)

  up_arrow = Button(arrows_frame, text="↑", command=lambda: swap_vars("up"), width="3")
  down_arrow = Button(arrows_frame, text="↓", command=lambda: swap_vars("down"), width="3")
  up_arrow.pack(side=TOP)
  down_arrow.pack(side=BOTTOM)

  label = Label(label_frame, text="Presentation path", bg=background, font="Helvetica 10 bold")
  label.pack(side=LEFT)

  entry = Entry(entry_frame, textvariable=var, width="30")
  def browse_dialog():
    filename = filedialog.askopenfilename(title="Select a presentation", filetype=[("PowerPoint files", "*.pptx")])
    var.set(filename)
    entry.xview_moveto(1)
  
  browse_btn = Button(entry_frame, text="Browse", command=browse_dialog)
  entry.pack(side=LEFT)

  # If the input is deletable, create a button to do that and pack it to the right, after the browse button.
  if deletable:
    def delete():
      container_frame.pack_forget()
      container_frame.destroy()

    delete_btn = Button(entry_frame, text="X", command=delete)
    delete_btn.pack(side=RIGHT)
  
  browse_btn.pack(side=RIGHT)
  return entry

def print_info(vars):
  for var in vars:
    print(var.get())

def add_input(frame, deletable=True):
  var = StringVar()
  entry = input_layout(frame, var, deletable=deletable)
  vars.append(var)
  presentation_entries.append(entry)

def check_file(filepath, exists=True, extention=".pptx"):
  if not exists:
    return os.path.exists(os.path.split(filepath)[0]) and not os.path.exists(filepath) and os.path.splitext(filepath)[1] == extention
  else:
    return os.path.exists(filepath) == exists and os.path.splitext(filepath)[1] == extention

def check_vars():
  for var in vars:
    if var.get() == "": return False
    if not check_file(var.get(), exists=True, extention=".pptx"): return False
  
  return True

def merge(output_entry, message_label):
  output_path = output_entry.get()
  vars_check = check_vars()
  output_check = check_file(output_path, exists=False, extention=".pptx")
  if vars_check and output_check:
    presentations = [var.get() for var in vars]
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
canvas = Canvas(canvas_frame)
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

# Button to add a new presentation to merge
add_btn = Button(add_btn_frame, text="Add presentation", command=lambda: add_input(input_frame), bg="#d04424", fg="white")
add_btn.pack(side=LEFT, pady="10", padx="5")

def browse_many_dialog():
  filenames = filedialog.askopenfilenames(initialdir="/", title="Select presentations", filetype=[("PowerPoint files", "*.pptx")])
  if len(vars) == 2 and vars[0].get() == "" and vars[1].get() == "" and len(filenames) >= 2:
    for i in range(len(filenames)):
      if i < 2:
        vars[i].set(filenames[i])
        presentation_entries[i].xview_moveto(1)
      else:
        add_input(input_frame)
        vars[-1].set(filenames[i])
        presentation_entries[-1].xview_moveto(1)
  else:
    for fname in filenames:
      add_input(input_frame)
      vars[-1].set(fname)
      presentation_entries[-1].xview_moveto(1)
  

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
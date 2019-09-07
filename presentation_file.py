import os
from tkinter import *
from tkinter import filedialog

class FileForm:
  def __init__(self, frame, is_input=True, text="Select a file", deletable=True, swappable=False,
               delete_callback=None, swap_callback=None,
               bg="SystemButtonFace", fg="black", filetype=[("PowerPoint files", "*.pptx")]):
    self.filetype = filetype
    self.is_input = is_input
    self.delete_callback = delete_callback
    assert not (swappable and swap_callback is None), "If PresentationFile is swappable it neets a swap_callback function!"
    self.swap_callback = swap_callback
    self.text = text
    self.deletable = deletable
    self.swappable = swappable
    self.frame = frame
    self.var = StringVar()
    self.bg = bg
    self.fg = fg
    self.container_frame = Frame(self.frame, bg=self.bg, pady="5", padx="5")
    self.container_frame.pack(side=TOP, fill=X, pady="5", padx="5")
    
    self.label_frame = Frame(self.container_frame, bg=self.bg)
    self.label_frame.pack(side=TOP, fill=X)
    
    self.entry_frame = Frame(self.container_frame, bg=self.bg)
    self.entry_frame.pack(side=TOP, fill=X)
    
    if self.swappable:
      self.arrows_frame = Frame(self.entry_frame, bg=self.bg)
      self.arrows_frame.pack(side=LEFT, fill=Y, padx="5")

      self.up_arrow = Button(self.arrows_frame, text="↑", command=lambda: self.swap_callback(self, "up"), width="3")
      self.down_arrow = Button(self.arrows_frame, text="↓", command=lambda: self.swap_callback(self, "down"), width="3")
      self.up_arrow.pack(side=TOP)
      self.down_arrow.pack(side=BOTTOM)

    self.label = Label(self.label_frame, text=self.text, bg=self.bg, fg=self.fg, font="Helvetica 10 bold")
    self.label.pack(side=LEFT)

    self.entry = Entry(self.entry_frame, textvariable=self.var, width="30", fg=self.fg)

    self.browse_btn = Button(self.entry_frame, text="Browse", bg=self.bg, fg=self.fg, command=self.browse_dialog)
    self.entry.pack(side=LEFT)

    # If the input is deletable, create a button to do that and pack it to the right, after the browse button.
    if self.deletable:
      self.delete_btn = Button(self.entry_frame, text="X", bg=self.bg, fg=self.fg, command=self.delete)
      self.delete_btn.pack(side=RIGHT)
    
    self.browse_btn.pack(side=RIGHT)

  @property
  def type(self):
    return "input"  if self.is_input else "output"
  
  @property
  def file_ext(self):
    # Takes the firs filetype, and uses the part after the * of the extention.
    # Extention has to be in the format '*.ext', otherwise it won't work!
    return self.filetype[0][1][1:]

  def set(self, value):
    self.var.set(value)
    self.entry.xview_moveto(1)
  
  def get(self):
    return self.var.get()

  def browse_dialog(self):
    if self.is_input:
      filename = filedialog.askopenfilename(title="Select a presentation", filetype=self.filetype)
      self.set(filename)
    else:
      filename = filedialog.asksaveasfilename(title="Save as", filetype=self.filetype)
      self.set(filename if os.path.splitext(filename)[1] == self.file_ext else filename + self.file_ext)
  
  def delete(self):
    self.container_frame.pack_forget()
    self.container_frame.destroy()
    # Try to use the callback if there is any
    try:
      self.delete_callback(self)
    except:
      pass

  def config(self, bg=None, fg=None):
    self.container_frame.config(bg=bg, fg=fg)
    self.label_frame.config(bg=bg, fg=fg)
    self.entry_frame.config(bg=bg, fg=fg)
    self.label.config(bg=bg, fg=fg)
    self.bg = bg if bg is not None else self.bg
    self.bg = fg if fg is not None else self.fg
  
  def file_errors(self):
    filepath = self.get()
    if not self.is_input:
      if not os.path.exists(os.path.split(filepath)[0]):
        return "Attempting to save at a location that was not found: {}".format(filepath)
      elif os.path.exists(filepath):
        return "File already exists. Replace is not possible."
      elif not os.path.splitext(filepath)[1] == self.file_ext:
        return "File {} does not have correct file extention. Should be '{}' but is '{}'".format(
            os.path.split(filepath)[1],
            self.file_ext,
            os.path.splitext(filepath)[1]
          )
      else:
        return None
    else:
      if not os.path.exists(filepath):
        return "File {} does not exist".format(filepath)
      elif not os.path.splitext(filepath)[1] == self.file_ext:
        return "File does not have correct file extention. Should be '{}' but is '{}'".format(self.file_ext, os.path.splitext(filepath)[1])
      else:
        return None


class FileFormList:
  def __init__(self, frame, is_input=True, inputs=[]):
    self.is_input = is_input
    self.inputs = inputs
    self.frame = frame

  @property
  def type(self):
    return "input"  if self.is_input else "output"

  def add_input(self, text="Presentation file", deletable=True, swappable=True):
    bg = "lightgrey" if (len(self.inputs) + 1) % 2 == 0 else "SystemButtonFace"
    self.append(FileForm(self.frame, is_input=self.is_input, text=text, bg=bg, deletable=deletable, swappable=swappable,
      delete_callback=self.delete, swap_callback=self.swap_inputs))

  def get_filepaths(self):
    return [obj.get() for obj in self.inputs]

  def get_value(self, index):
    return self.inputs[index].get()

  def set_value(self, index, value):
    self.inputs[index].set(value)
  
  def append(self, obj):
    self.inputs.append(obj)
  
  def pop(self, index):
    self.inputs.pop(index)
    for i, input in enumerate(self.inputs):
      background = "lightgrey" if (i + 1) % 2 == 0 else "SystemButtonFace"
      input.config(bg=background)
  
  def find(self, obj):
    return [id(i) for i in self.inputs].index(id(obj))

  def swap_inputs(self, prs_file, direction):
    index = self.find(prs_file)
    if direction == "up" and index - 1 >= 0:
      buff = prs_file.get()
      prs_file.var.set(self.get_value(index - 1))
      self.set_value(index - 1, buff)
    elif direction == "down" and index + 1 < len(self.inputs):
      buff = prs_file.get()
      prs_file.set(self.get_value(index + 1))
      self.set_value(index + 1, buff)
  
  def delete(self, prs_file):
    self.pop(self.find(prs_file))
  
  def input_errors(self):
    errors = []
    for input in self.inputs:
      err = input.file_errors()
      if err is not None: errors.append(err)
    
    return errors
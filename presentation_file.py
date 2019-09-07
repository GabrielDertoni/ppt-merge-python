import os
from tkinter import *
from tkinter import filedialog

class PresentationFile:
  # Static variables and methods
  inputs = []

  @staticmethod
  def get_filepaths():
    return [obj.get() for obj in PresentationFile.inputs]

  @staticmethod
  def get_value(index):
    return PresentationFile.inputs[index].get()

  @staticmethod
  def set_value(index, value):
    PresentationFile.inputs[index].set(value)
  
  @staticmethod
  def append(obj):
    PresentationFile.inputs.append(obj)
  
  @staticmethod
  def pop(index):
    PresentationFile.inputs.pop(index)
    for input in PresentationFile.inputs:
      input.update()
  
  @staticmethod
  def find(obj):
    return [id(i) for i in PresentationFile.inputs].index(id(obj))
  
  @staticmethod
  def check():
    for input in PresentationFile.inputs:
      if input.get() == "": return False
      if not input.check_file(exists=True, extention=".pptx"): return False
    
    return True

  def __init__(self, frame, text="Presentation file", deletable=True, swappable=True,
               delete_callback=None, swap_callback=None):
    self.delete_callback = delete_callback
    self.swap_callback = swap_callback
    self.text = text
    self.deletable = deletable
    self.swappable = swappable
    self.frame = frame
    self.var = StringVar()
    self.index = len(PresentationFile.inputs)

    self.background = "lightgrey" if (self.index + 1) % 2 == 0 else "SystemButtonFace"
    
    self.container_frame = Frame(self.frame, bg=self.background, pady="5", padx="5")
    self.container_frame.pack(side=TOP, fill=X, pady="5", padx="5")
    
    self.label_frame = Frame(self.container_frame, bg=self.background)
    self.label_frame.pack(side=TOP, fill=X)
    
    self.entry_frame = Frame(self.container_frame, bg=self.background)
    self.entry_frame.pack(side=TOP, fill=X)
    
    if self.swappable:
      self.arrows_frame = Frame(self.entry_frame, bg=self.background)
      self.arrows_frame.pack(side=LEFT, fill=Y, padx="5")

      self.up_arrow = Button(self.arrows_frame, text="↑", command=lambda: self.swap_vars("up"), width="3")
      self.down_arrow = Button(self.arrows_frame, text="↓", command=lambda: self.swap_vars("down"), width="3")
      self.up_arrow.pack(side=TOP)
      self.down_arrow.pack(side=BOTTOM)

    self.label = Label(self.label_frame, text=self.text, bg=self.background, font="Helvetica 10 bold")
    self.label.pack(side=LEFT)

    self.entry = Entry(self.entry_frame, textvariable=self.var, width="30")

    self.browse_btn = Button(self.entry_frame, text="Browse", command=self.browse_dialog)
    self.entry.pack(side=LEFT)

    # If the input is deletable, create a button to do that and pack it to the right, after the browse button.
    if self.deletable:
      self.delete_btn = Button(self.entry_frame, text="X", command=self.delete)
      self.delete_btn.pack(side=RIGHT)
    
    self.browse_btn.pack(side=RIGHT)
    
    PresentationFile.append(self)

  def set(self, value):
    self.var.set(value)
    self.entry.xview_moveto(1)
  
  def get(self):
    return self.var.get()

  def browse_dialog(self):
    filename = filedialog.askopenfilename(title="Select a presentation", filetype=[("PowerPoint files", "*.pptx")])
    self.set(filename)
  
  def swap_vars(self, direction):
    if direction == "up" and self.index - 1 >= 0:
      buff = self.get()
      self.var.set(PresentationFile.get_value(self.index - 1))
      PresentationFile.set_value(self.index - 1, buff)
    elif direction == "down" and self.index + 1 < len(PresentationFile.inputs):
      buff = self.get()
      self.set(PresentationFile.get_value(self.index + 1))
      PresentationFile.set_value(self.index + 1, buff)
  
  def delete(self):
    self.container_frame.pack_forget()
    self.container_frame.destroy()
    # Try to use the callback if there is any
    # try:
    #   self.delete_callback(self)
    # except:
    #   pass
    PresentationFile.pop(self.index)

  def update(self):
    self.index = PresentationFile.find(self)
    self.background = "lightgrey" if (self.index + 1) % 2 == 0 else "SystemButtonFace"
    self.container_frame.config(bg=self.background)
    self.label_frame.config(bg=self.background)
    self.entry_frame.config(bg=self.background)
    self.label.config(bg=self.background)
  
  def check_file(self, exists=True, extention=".pptx"):
    filepath = self.get()
    if not exists:
      return os.path.exists(os.path.split(filepath)[0]) and not os.path.exists(filepath) and os.path.splitext(filepath)[1] == extention
    else:
      return os.path.exists(filepath) == exists and os.path.splitext(filepath)[1] == extention


# class PresentationFileList:
#   def __init__(self, inputs=[]):
#     self.inputs = inputs

#   def add_input(self):
#     self.append(PresentationFile())

#   def get_filepaths(self):
#     return [obj.get() for obj in self.inputs]

#   def get_value(self, index):
#     return self.inputs[index].get()

#   def set_value(self, index, value):
#     self.inputs[index].set(value)
  
#   def append(self, obj):
#     self.inputs.append(obj)
  
#   def pop(self, index):
#     self.inputs.pop(index)
#     for input in self.inputs:
#       input.update()
  
#   def find(self, obj):
#     return [id(i) for i in self.inputs].index(id(obj))

#   def swap_vars(self, direction):
#     if direction == "up" and self.index - 1 >= 0:
#       buff = self.get()
#       self.var.set(self.get_value(self.index - 1))
#       self.set_value(self.index - 1, buff)
#     elif direction == "down" and self.index + 1 < len(self.inputs):
#       buff = self.get()
#       self.set(self.get_value(self.index + 1))
#       self.set_value(self.index + 1, buff)
  
#   def delete(self, index):
#     self.container_frame.pack_forget()
#     self.container_frame.destroy()
#     self.pop(self.index)
  
#   def update(self):
#     self.index = self.find(self)
#     self.background = "lightgrey" if (self.index + 1) % 2 == 0 else "SystemButtonFace"
#     self.container_frame.config(bg=self.background)
#     self.label_frame.config(bg=self.background)
#     self.entry_frame.config(bg=self.background)
#     self.label.config(bg=self.background)
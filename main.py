import os
import win32com.client
import argparse

def merge_presentations(presentations, path):
  ppt_instance = win32com.client.Dispatch('PowerPoint.Application')
  # open the powerpoint presentation headless in background
  prs = ppt_instance.Presentations.open(os.path.abspath(presentations[0]), True, False, False)

  for i in range(1, len(presentations)):
    prs.Slides.InsertFromFile(os.path.abspath(presentations[i]), prs.Slides.Count)

  prs.SaveAs(os.path.abspath(path))
  prs.Close()

  #kills ppt_instance
  ppt_instance.Quit()
  del ppt_instance


def parse_args():
  parser = argparse.ArgumentParser(description="Merge a list of powerpoint presentations")
  parser.add_argument('presentations', type=str, nargs='+',
    help="""
    A list of absolute or relative paths to PowerPoint presentations (.pptx).
    Needs at least two presentations to merge
    """)
  parser.add_argument('--name', type=str, default="merge.pptx", help="The path to save file in. Default is './merge.pptx'")
  return parser.parse_args()


def main(args):
  presentations = args.presentations
  assert os.path.splitext(args.name)[1] == ".pptx", "Output file must be of extention .pptx"
  assert len(presentations) >= 2, "The must be at least thwo files to merge."
  for prs in presentations:
    assert os.path.splitext(prs)[1] == ".pptx", "All files must be of extention .pptx"
    assert os.path.exists(prs), "'{}' was not found or does not exist.".format(prs)
  
  merge_presentations(presentations, args.name)

if __name__ == '__main__':
  main(parse_args())
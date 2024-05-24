import cv2
import pandas as pd
import xlrd
import numpy as np
from PIL import Image, ImageDraw, ImageFont
import itertools

# Load the font
FONT_SIZE = 50
FONT_PATH = '/font/DejaVuSans.ttf'
FONT_FACE = ImageFont.truetype(FONT_PATH, FONT_SIZE)

def getInputTextSize(text):
  (text_width, text_height) = FONT_FACE.getsize(text)
  return (text_width, text_height)

def splitList(input_list, line_number):
  length_list = len(input_list)
  el_per_line = length_list // line_number
  split_list = []
  for ind in range(line_number):
    split_list.append(input_list[0:el_per_line])
    del input_list[0:el_per_line]
  return split_list

def getListText(text, frame_width):
  (text_width, text_height) = getInputTextSize(text)
  if text_width < frame_width:
    return [text]
  
  line_text = text_width // frame_width
  remainder = text_width % frame_width
  if (remainder > 0):
    line_text = line_text + 1

  split_text = text.split(" ")
  split_list_text = splitList(split_text, line_text)
  list_text = []
  for item in split_list_text:
    list_text.append(" ".join(item))

  return list_text

def inputText(draw, text, output_video_size, list_text):
  for index in range(len(list_text)):
    text = list_text[index]
    (text_width, text_height) = getInputTextSize(text)
    text_position = (int((output_video_size[0] - text_width) // 2), int((index + 1) * (text_height + 10)))
    draw.text(text_position, text, font=FONT_FACE, fill=(0, 0, 0))

def drawBox(draw, text_position, list_text, video_size):
  for index in range(len(list_text)):
    (text_width, text_height) = getInputTextSize(list_text[index])
    background_position = (int((video_size[0] - text_width) // 2), int(text_position[1] + (index + 1) * (text_height + 15)))
    background_size = (int(background_position[0] + text_width), int(background_position[1] - text_height - 20))
    draw.rectangle([background_position, background_size], fill=(255, 255, 255))

def addTextToVideo(url, text):
  print(">>> Inserting '" + text +"' to: " + url + " ...")
  video_capture = cv2.VideoCapture(url)

  # Get video properties
  fps = video_capture.get(cv2.CAP_PROP_FPS)
  frame_width = int(video_capture.get(cv2.CAP_PROP_FRAME_WIDTH))
  frame_height = int(video_capture.get(cv2.CAP_PROP_FRAME_HEIGHT))

  # Create VideoWriter object to save the output
  fourcc = cv2.VideoWriter_fourcc(*'mp4v')
  output_video_size = (frame_width, frame_height)
  output_video_name = url.split("/").pop()
  output_video_path = './output/result_' + output_video_name
  output_video = cv2.VideoWriter(output_video_path, fourcc, fps, (frame_width, frame_height))

  # Specify the path to the downloaded font file
  font_path = '/DejaVuSans.ttf'
  
  (text_width, text_height) = getInputTextSize(text)
  input_list_text = getListText(text, frame_width)
  text_position = (0, text_height)

  # Process each frame and add text caption
  while True:
    ret, frame = video_capture.read()
    if not ret:
        break

    # Convert the frame to RGB (PIL uses RGB)
    frame_pil = Image.fromarray(cv2.cvtColor(frame, cv2.COLOR_BGR2RGB))

    # Add white background
    draw = ImageDraw.Draw(frame_pil)
    drawBox(draw, text_position, input_list_text, output_video_size)
    
    # Add black text    
    inputText(draw, text, output_video_size, input_list_text)

    # Convert the frame back to BGR
    frame = cv2.cvtColor(np.array(frame_pil), cv2.COLOR_RGB2BGR)

    # Write the frame to the output video
    output_video.write(frame)

  # Release video capture and writer
  video_capture.release()
  output_video.release()
  print("Done!")
  cv2.destroyAllWindows()

# Read data from excel file (xlsx)
INPUT_DATA_PATH = 'input.xlsx'
def readInput():
  data = pd.read_excel(INPUT_DATA_PATH, sheet_name="Sheet1", index_col=[0,1])
  if data.index.size == 0:
    print("Empty data!")
  return data.index

# main
inputData = readInput()
for data in inputData:
  url=data[0]
  text=data[1]
  addTextToVideo(url, text)

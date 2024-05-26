import cv2
import pandas as pd
import xlrd
import numpy as np
from PIL import Image, ImageDraw, ImageFont
import itertools
import datetime
import pathlib
import threading
import random

# Load the font
FONT_SIZE = 50
FONT_PATH = './font/Roboto/Roboto-Bold.ttf'
FONT_FACE = ImageFont.truetype(FONT_PATH, FONT_SIZE)
INPUT_DATA_PATH = 'input.xlsx'
OUTPUT_FOLDER = f'./output/{datetime.datetime.now().strftime("%Y-%m-%d %H-%M-%S")}/'

# Thread props
TOTAL_THREADS = 10
THREAD_COUNTER = 0
THREAD_FREE = True


# create output folder
pathlib.Path(OUTPUT_FOLDER).mkdir(parents=True, exist_ok=True)

def get_input_text_size(text):
  (text_width, text_height) = FONT_FACE.getsize(text)
  return (text_width, text_height)

def split_list(input_list, line_number):
  length_list = len(input_list)
  el_per_line = length_list // line_number
  split_list = []
  for ind in range(line_number):
    split_list.append(input_list[0:el_per_line])
    del input_list[0:el_per_line]
  return split_list

def get_list_text(text, frame_width):
  (text_width, text_height) = get_input_text_size(text)
  if text_width < frame_width:
    return [text]

  line_text = text_width // frame_width
  remainder = text_width % frame_width
  if (remainder > 0):
    line_text = line_text + 1

  split_texts = text.split(" ")
  split_list_text = split_list(split_texts, line_text)
  list_text = []
  for item in split_list_text:
    list_text.append(" ".join(item))

  return list_text

def input_text(draw, text, output_video_size, list_text):
  for index in range(len(list_text)):
    text = list_text[index]
    (text_width, text_height) = get_input_text_size(text)
    text_position = (int((output_video_size[0] - text_width) // 2), int((index + 1) * text_height))
    draw.text(text_position, text, font=FONT_FACE, fill=(0, 0, 0))

def draw_box(draw, text_position, list_text, output_video_size):
  for index in range(len(list_text)):
    (text_width, text_height) = get_input_text_size(list_text[index])
    background_position = (int((output_video_size[0] - text_width) // 2 - 10), int((index + 1) * text_height))
    background_size = (int(background_position[0] + text_width + 20), int(background_position[1] + text_height))
    draw.rounded_rectangle((background_position, background_size), 20, fill=(255, 255, 255))

def add_text_to_video(url, text):
  global THREAD_FREE, THREAD_COUNTER
  print(">>> Inserting '" + text +"' to: " + url + " ...")
  video_capture = cv2.VideoCapture(url)

  # Get video properties
  fps = video_capture.get(cv2.CAP_PROP_FPS)
  frame_width = int(video_capture.get(cv2.CAP_PROP_FRAME_WIDTH))
  frame_height = int(video_capture.get(cv2.CAP_PROP_FRAME_HEIGHT))

  if int(fps) == 0 | frame_width == 0 | frame_height == 0:
    print("Url: " + url + " is not found!")
    THREAD_FREE = True
    THREAD_COUNTER = THREAD_COUNTER - 1
    return

  # Create VideoWriter object to save the output
  fourcc = cv2.VideoWriter_fourcc(*'mp4v')
  output_video_size = (frame_width, frame_height)
  output_video_name = url.split("/").pop()
  output_video_path = OUTPUT_FOLDER + output_video_name
  output_video = cv2.VideoWriter(output_video_path, fourcc, fps, (frame_width, frame_height))

  (text_width, text_height) = get_input_text_size(text)
  input_list_text = get_list_text(text, frame_width)
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
    draw_box(draw, text_position, input_list_text, output_video_size)

    # Add black text
    input_text(draw, text, output_video_size, input_list_text)

    # Convert the frame back to BGR
    frame = cv2.cvtColor(np.array(frame_pil), cv2.COLOR_RGB2BGR)

    # Write the frame to the output video
    output_video.write(frame)

  # Release video capture and writer
  video_capture.release()
  output_video.release()
  print("Done!")
  
  THREAD_FREE = True
  THREAD_COUNTER = THREAD_COUNTER - 1
  cv2.destroyAllWindows()

def thread_work(url, text):
  global THREAD_COUNTER
  THREAD_COUNTER = THREAD_COUNTER + 1
  worker = threading.Thread(target=add_text_to_video, args=(url, text))
  worker.start()

# Read data from excel file (xlsx)
def read_input():
  data = pd.read_excel(INPUT_DATA_PATH, sheet_name="Sheet1", index_col=[0,1])
  if data.index.size == 0:
    print("Empty data!")
  return data.index

def convert_data(input_data):
  output_data = []
  for data in input_data:
    output_data.append([data[0], data[1], False])
  return output_data

def delete_data_element(list_data):
  return [x for x in list_data if x[2] == False]

# main
input_data = read_input()
list_data = convert_data(input_data)

while True:
  if len(list_data) == 0:
    break

  if (len(list_data) <= TOTAL_THREADS):
    for data in list_data:
      url = data[0]
      text = data[1]
      data[2] = True
      thread_work(url, text)
    break
  
  for data in list_data:
    if THREAD_COUNTER > TOTAL_THREADS - 1:
      THREAD_FREE = False
      list_data = delete_data_element(list_data)
      break

    if len(list_data) < TOTAL_THREADS & THREAD_COUNTER == 0:
      THREAD_FREE = True
      list_data = delete_data_element(list_data)

    if THREAD_FREE:
      url = data[0]
      text = data[1]
      data[2] = True
      thread_work(url, text)

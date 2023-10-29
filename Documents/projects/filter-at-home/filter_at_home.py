import win32ui
import win32api
import win32con
import keyboard
import pyautogui
from PIL import Image, ImageEnhance, ImageOps
from pytesseract import pytesseract
import time
import re
import pandas as pd
import numpy as np
from pprint import pprint
from IPython.display import clear_output

def get_screenshot(x,y,box='left'):
    destination = (x,y)
    pyautogui.moveTo(x, y, duration=0.01)
    # pyautogui.moveTo(x+15, duration=0.3)
    pyautogui.doubleClick() 
    time.sleep(0.075)
    t0 = time.time()
    im = pyautogui.screenshot()
    print(type(a))
    t1 = time.time()
    print("take screen:",t1-t0)
    # im = Image.open("cur_item.png")
    if box=='right':
        left = destination[0]+80
        top = destination[1]-800
        right = destination[0]+435
        bottom = destination[1]+100
    elif box=='left':
        left = destination[0]-400
        top = destination[1]-800
        right = destination[0]-45
        bottom = destination[1]+100
    
    # Cropped image of above dimension
    # (It will not change original image)
    im1 = im.crop((left, top, right, bottom))
    t2 = time.time()
    # print("crop:",t2-t1)
    # # enhancer = ImageEnhance.Brightness(im1)
    # # im2 = enhancer.enhance(10)
    # # (width, height) = (im2.width // 2, im2.height // 2)
    # # im_resized = im2.resize((width, height))
    # # display(im2)
    # # display(im_resized)
    return im1

def resize_image(img_var, target_width):
    wpercent = (basewidth/float(img_var.size[0]))
    hsize = int((float(img_var.size[1])*float(wpercent)))
    result = img_var.resize((basewidth,hsize), Image.Resampling.LANCZOS)
    return result



def yes_binary(img_var,threshold):
    t0 = time.time()
    #1 reduce color
    enhancer_c = ImageEnhance.Color(img_var)
    im2 = enhancer_c.enhance(0.1)
    # display(im2)
     
    #1 get more brightness
    enhancer_c = ImageEnhance.Brightness(im2)
    im3 = enhancer_c.enhance(3)
    # display(im3)
    
    # #3 sharpness(?)
    # enhancer_s = ImageEnhance.Sharpness(im3)
    # im4 = enhancer_s.enhance(2)
    # display(im4)
    
    # extract_text(im3)
    t1=time.time()
    print('Binarized in:',t1-t0)
    temp1 = im3.convert("L")
    # display(temp1)
    image_file = temp1.point( lambda p: 255 if p > threshold else 0 )
    result = image_file.convert('1')
    result = ImageOps.invert(result)
    text = extract_text(result)
    display(result)
    t1=time.time()
    print('Processed in:',t1-t0)
    return text

 def extract_text(img_var):
     # Defining paths to tesseract.exe
    # and the image we would be using
    path_to_tesseract = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
    t0=time.time()
    # Providing the tesseract executable
    # location to pytesseract library
    pytesseract.tesseract_cmd = path_to_tesseract

    # Passing the image object to image_to_string() function
    # This function will extract the text from the image
    custom_oem_psm_config = r'--psm 4'
    text = pytesseract.image_to_string(img_var,config=custom_oem_psm_config)
    text = text.replace('\n\n','\n')
    # print(text)
  
    # t1=time.time()
    # print("extract and process text:",t1-t0)
    
    #gererate_rows
    text = (re.sub('[^ 0-9a-zA-Z\+\-\[\]\'\n\'\%\*\,\.]', '',text))
    # print(text)
    row_split = text.split('\n')
    # print(row_split)
    text_clean = []
    find_start = 0
    for row in row_split:
        if row.lower().find('sacred')!= -1 and find_start==0:
            find_start+=1
        elif row.lower().find('ancestral')!= -1 and find_start==0:
            find_start+=1
        if row.lower().find('requires') != -1:
            break
        if find_start>0:
            text_clean.append(row)
    print('tess output:',text_clean)
    return(text_clean)

def evaluate_item_object(my_text, 
						 min_item_power, 
						 target_aspects, 
						 affect_unique=False, 
						 ignore_aspects=True, 
						 target_builds):
    #Remove empty strings (tesseract artifacts)
    print(my_text)
    clean_text = []
    print('count_str_raw:',len(my_text))
    for line in my_text:
        if re.sub('[^0-9a-zA-Z]', '',line) != '':
            clean_text.append(line)
    print('count_str_clean:',len(clean_text))
    print(clean_text)
    
    #check if text is present
    if len(clean_text)==0:
        # pyautogui.press('space')
        print(''' 
                  ----------------------
                  <<<Marked as trash >>> 
                  No text to parse
                  ----------------------''')
        return

    #Create item object    
    item = {}    
    #Get item power
    st=0
    print(clean_text)
    for index, line in enumerate(clean_text):
        if line.lower().find('item')!= -1:
            item['Item_Power'] = int(re.sub('[^0-9]', '',clean_text[index]))
            st = index+1
            break

    #Check for item power
    if item['Item_Power'] < min_item_power:
        pyautogui.press('space')
        print(''' 
                  ----------------------
                  <<<Marked as trash >>> 
                  Low Item Power
                  ----------------------''')
        return 

    #merge list of lines into single lowercase text
    affixes_raw = []
    for line in clean_text[st:]:
        affixes_raw.append(re.sub('[^a-zA-Z]', '',line)).lower()
    affixes_body = ' '.join(affixes_raw)

    
    desired_tiers = re.compile('(ancestral|sacred)')
    item['Item_Tier'] = desired_tiers.findall(affixes_body)[0]
    rarities = re.compile('(rare|legendary|unique)')
    item['Item_Rarity'] = rarities.findall(affixes_body)[0]

    #Check for useful aspect
    if item['Item_Rarity']=='legendary' and ignore_aspects== False:
        for asp_keyword in target_aspects:
        	tmp_pattern = re.compile(asp_keyword)
            if len(tmp_pattern.findall(affixes_body))>0:
                print(''' 
                      ----------------------
                      <<<Deemed good >>> 
                      Useful Aspect found!
                      ----------------------''')
                return
    
    #Check for unique:
    if item['Item_Rarity']=='unique' and affect_unique==False:
        print(''' 
                  ----------------------
                  <<<Deemed good >>> 
                  Unique found!
                  ----------------------''')
        return    
    


    #get all possible slots to look for 
    temp_slots = []
	for build in target_builds:
	    for k in build.keys():
	        temp_slots.append(k)
	slots = list(set(temp_slots))
	print(slots)
	#Get item slot
	tmp_pattern = re.compile('('+'|'.join(slots)+')')

	item['Item_Slot'] = tmp_pattern.findall(affixes_body)[0]
    
    
      
    
    #Check for item power
    if item['Item_Power'] < min_item_power:
        pyautogui.press('space')
        print(''' 
                  ----------------------
                  <<<Marked as trash >>> 
                  Low Item Power
                  ----------------------''')
        return 
    

    
    #find keywords within text, if found < minimum, the item is salvaged
    min_affixes = target_builds[item['Item_Slot']][1]
    print(min_affixes)
    keywords_found_count = 0
    keywords_found = []
    for keyword in target[item['Item_Slot']][0]:
        if affixes_body.lower().find(keyword) != -1:
            keywords_found_count += 1
            keywords_found.append(keyword)
    if keywords_found_count < min_affixes:
        pyautogui.press('space')
        print(''' 
              ----------------------
              <<<Marked as trash >>> 
              Minimum of {} target affixes required,
              Target affixes list:
              {}
              {} found on the item:
              {}
              ----------------------'''.format(min_affixes,target[item['Item_Slot']][0],keywords_found_count,keywords_found))
    else:
        print(''' 
              ----------------------
              <<<Deemed good >>> 
              Minimum of {} target affixes required,
              Target affixes list:
              {}
              {} found on the item:
              {}
              ----------------------'''.format(min_affixes,target[item['Item_Slot']][0],keywords_found_count,keywords_found))
    return    
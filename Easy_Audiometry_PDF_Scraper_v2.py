#!/usr/bin/env python
# coding: utf-8

# # Programma voor de extractie van audiometrische data uit PDF EasyViewer

# Eigenaar: Guido Cattani

# Datum van deze versie: 20 December 2023 

# In[1]:


from colorama import Fore, Style, init


# In[2]:


init()
a ='\n\tWelkom bij EasyPDF AudioScraper versie 2 (20 december 2023).' 
b ='\nDit programma kan een Easy Data pdf-file met toon- en spraakaudiogram analyseren en een beschrijving maken.'
c ='\nDe beschrijving wordt opgeslagen in een textfile in uw (Onedrive) downloadlocatie.'
d = '\nHet programma is niet altijd nauwkeurig. Controleer daarom altijd goed de beschrijving.'
e = '\n© 2023 Guido Cattani'
print(Fore.WHITE + Style.BRIGHT + a + b + c + d + e + '\n' + Style.RESET_ALL)


# In[3]:


# Import libraries
import re
import os
from os.path import expanduser, join
import fitz  # PyMuPDF
from PIL import Image
import numpy as np


# In[4]:


# Constants of this program (for maintenance)

# Define constants for rectangle coordinates, depends heavily from pdf layout

# Coordinates Header
HEADER_LEFT = 0
HEADER_TOP = 0
HEADER_RIGHT = 600
HEADER_BOTTOM = 140

# Coordinates Audiogram Left ear
LEFT_EAR_LEFT = 321
LEFT_EAR_TOP = 140
LEFT_EAR_RIGHT = 600
LEFT_EAR_BOTTOM = 449

# Coordinates Audiogram Right ear
RIGHT_EAR_LEFT = 0
RIGHT_EAR_TOP = 140
RIGHT_EAR_RIGHT = 320
RIGHT_EAR_BOTTOM = 449

# Coordinates Bottom Text
BOTTOM_LEFT = 0
BOTTOM_TOP = 660
BOTTOM_RIGHT = 600
BOTTOM_BOTTOM = 800

# Coordinates Speech Audiogram
SPEECH_LEFT = 40
SPEECH_TOP = 460
SPEECH_RIGHT = 590
SPEECH_BOTTOM = 650

# Coordinates Speech Audiogram Right Ear (RE)
SPEECH_RE_LEFT = 40
SPEECH_RE_TOP = 460
SPEECH_RE_RIGHT = 330
SPEECH_RE_BOTTOM = 650

# Coordinates Speech Audiogram Left Ear (LE)
SPEECH_LE_LEFT = 314
SPEECH_LE_TOP = 460
SPEECH_LE_RIGHT = 590
SPEECH_LE_BOTTOM = 650

# Coordinates to determin speech-audiogram origin and scale
SCORE_100 = 19
SCORE_50 = 68
SCORE_0 = 118
LEVEL_0_RE = 44
LEVEL_120_RE = 262
LEVEL_0_LE = 26
LEVEL_120_LE = 243


# Tuple containing the speech audiogram origin coordinates (x, y)
SPEECH_RE_ORIGIN = (LEVEL_0_RE, SCORE_0) # Origin of right ear (RE) speech-audiogram
SPEECH_LE_ORIGIN = (LEVEL_0_LE, SCORE_0) # Origin of right ear (LE) speech-audiogram


# Define RGB range for colors to detect in speech-audiogram
# Define RGB range for red color
RED_COLOR_RANGE = [(217, 256), (100, 160), (100, 160)]

# Define RGB range for blue color
BLUE_COLOR_RANGE = [(100, 160), (100, 160), (190, 256)]


# In[5]:


# Define constants for hearing loss extent thresholds
VERY_SEVERE_EXTENT = 80
SEVERE_TO_VERY_SEVERE_EXTENT = 75
SEVERE_EXTENT = 60
MODERATE_TO_SEVERE_EXTENT = 55
MODERATE_EXTENT = 40
LIGHT_TO_MODERATE_EXTENT = 35
LIGHT_EXTENT = 20


# In[6]:


def open_pdf_safely(pdf_path):
    """
    Safely opens a PDF document using PyMuPDF.

    Args:
        pdf_path (str): Path to the PDF file.

    Returns:
        fitz.Document or None: A PyMuPDF document object if the file is found, else None.
    """
    try:
        pdf_document = fitz.open(pdf_path)
        return pdf_document
    except FileNotFoundError:
        print(f"Bestand niet gevonden: {pdf_path}")
        return None
    except fitz.PyMuPDFError as e:
        print(f"Een fout is opgetreden tijdens het openen van het PDF-bestand: {e}")
        return None
    except Exception as e:
        print(f"Een onbekende fout is opgetreden tijdens het openen van het PDF-bestand: {e}")
        return None


# In[7]:


def extract_text_from_rect(pdf_path, rect):
    """
    Extracts text from a specified rectangle in a PDF document using PyMuPDF.

    Args:
        pdf_path (str): Path to the PDF file.
        rect (fitz.Rect): Rectangle defining the area to extract text from.

    Returns:
        str: Extracted text from the specified rectangle.
    """
    try:
        pdf_document = fitz.open(pdf_path)
        page = pdf_document[0]  # Assuming 1 page PDF
        text = page.get_text("text", clip=rect)
        pdf_document.close()
        return text
    except FileNotFoundError:
        print(f"Bestand niet gevonden: {pdf_path}")
        return ""
    except fitz.PyMuPDFError as e:
        print(f"Een fout is opgetreden bij het extraheren van tekst uit het PDF-bestand: {e}")
        return ""
    except Exception as e:
        print(f"Een onbekende fout is opgetreden bij het extraheren van tekst uit het PDF-bestand: {e}")
        return ""


# In[8]:


def extract_header_text(pdf_infile):
    """
    Extracts header text (with patientID) from a PDF document using PyMuPDF.

    Args:
        pdf_infile (str): Path to the PDF file.

    Returns:
        str: Extracted header text.
    """
    try:
        header_rect = fitz.Rect(HEADER_LEFT, HEADER_TOP, HEADER_RIGHT, HEADER_BOTTOM)
        header_text = extract_text_from_rect(pdf_infile, header_rect)
        return header_text
    except fitz.PyMuPDFError as e:
        print(f"An error occurred while extracting text from the header: {e}")
        return ""
    except Exception as e:
        print(f"An unknown error occurred while extracting text from the header: {e}")
        return ""


# In[9]:


def extract_right_audiogram_text(pdf_infile):
    """
    Extracts text from the right ear audiogram section of a PDF audiogram document using PyMuPDF.

    Args:
        pdf_infile (str): Path to the PDF file.

    Returns:
        str: Extracted text from the right ear audiogram section.
    """
    try:
        right_ear_audiogram_rect = fitz.Rect(RIGHT_EAR_LEFT, RIGHT_EAR_TOP, RIGHT_EAR_RIGHT, RIGHT_EAR_BOTTOM)
        right_ear_audiogram_text = extract_text_from_rect(pdf_infile, right_ear_audiogram_rect)
        return right_ear_audiogram_text
    except fitz.PyMuPDFError as e:
        print(f"An error occurred while extracting text from the right ear audiogram section: {e}")
        return ""
    except Exception as e:
        print(f"An unknown error occurred while extracting text from the right ear audiogram section: {e}")
        return ""


# In[10]:


def extract_bottom_text(pdf_infile):
    """
    Extracts bottom text from a PDF document using PyMuPDF.

    Args:
        pdf_infile (str): Path to the PDF file.

    Returns:
        str: Extracted bottom text.
    """
    try:
        bottom_rect = fitz.Rect(BOTTOM_LEFT, BOTTOM_TOP, BOTTOM_RIGHT, BOTTOM_BOTTOM)
        bottom_text = extract_text_from_rect(pdf_infile, bottom_rect)
        return bottom_text
    except fitz.PyMuPDFError as e:
        print(f"An error occurred while extracting text from the bottom of the PDF: {e}")
        return ""
    except Exception as e:
        print(f"An unknown error occurred while extracting text from the bottom of the PDF: {e}")
        return ""


# In[11]:


def extract_left_audiogram_text(pdf_infile):
    """
    Extracts text from the left ear audiogram section of a PDF audiogram document using PyMuPDF.

    Args:
        pdf_infile (str): Path to the PDF file.

    Returns:
        str: Extracted text from the left ear audiogram section.
    """
    try:
        left_ear_audiogram_rect = fitz.Rect(LEFT_EAR_LEFT, LEFT_EAR_TOP, LEFT_EAR_RIGHT, LEFT_EAR_BOTTOM)
        left_ear_audiogram_text = extract_text_from_rect(pdf_infile, left_ear_audiogram_rect)
        return left_ear_audiogram_text
    except fitz.PyMuPDFError as e:
        print(f"An error occurred while extracting text from the left ear audiogram section: {e}")
        return ""
    except Exception as e:
        print(f"An unknown error occurred while extracting text from the left ear audiogram section: {e}")
        return ""


# In[12]:


def provide_hearing_loss_data(audiogram_text):
    """
    Extracts hearing loss data from audiogram text looking for the strings 'Be:', 'FI:' and 'FIH:'.

    Args:
        audiogram_text (str): Text containing audiogram data.

    Returns:
        tuple: Tuple containing hearing loss data as integers (BG_05_1, LG_05_1, LG_1_4, ABG).
    """
    try:
        lines = audiogram_text.split('\n')

        # Define the patterns to search for
        patterns = ['Be:', 'FI:', 'FIH:']

        # Initialize the variables with None
        BG_05_1 = None  # Average Bone Conduction Thresholds [dBHL] at 0.5, 1, 2 kHz
        LG_05_1 = None  # Average Air Conduction Thresholds [dBHL] at 0.5, 1, 2 kHz
        LG_1_4 = None   # Average Air Conduction Thresholds [dBHL] at 1, 2, 4 kHz
        ABG = None      # Average Air-Bone Gap [dBHL] at 1, 2, 4 kHz (LG_05_1 - BG_05_1)

        # Traverse lines to find hearing loss data
        for line in lines:
            for pattern in patterns:
                if pattern in line:
                    try:
                        value_str = line.split(pattern)[1].strip()
                        if value_str == "-":
                            value = None
                        else:
                            value = int(value_str)  # Convert the extracted string to an integer
                    except (IndexError, ValueError):
                        value = None  # Set to None when extraction fails

                    if pattern == 'Be:':
                        BG_05_1 = value
                    elif pattern == 'FI:':
                        LG_05_1 = value
                    elif pattern == 'FIH:':
                        LG_1_4 = value
                    break  # Break the loop after finding the first occurrence

        if BG_05_1 is not None and LG_05_1 is not None and BG_05_1:
            ABG = LG_05_1 - BG_05_1

        return BG_05_1, LG_05_1, LG_1_4, ABG

    except Exception as e:
        print(f"An error occurred while extracting hearing loss data: {e}")
        return None, None, None, None


# In[13]:


def type_hearing_loss(hearing_loss_data):
    """
    Determines the type of hearing loss based on provided data, providing a description.

    Args:
        hearing_loss_data (tuple): Tuple containing hearing loss data (BG_05_1, LG_05_1, LG_1_4, ABG).

    Returns:
        str: Description of hearing loss type.
    """
    BG_05_1, LG_05_1, LG_1_4, ABG = hearing_loss_data
    after_extent_hl = ' dB (gemiddeld bij 1, 2 en 4 kHz)'
    
    if LG_1_4 == None: type_hearing_loss = 'gemengd/perceptief gehoorverlies, met een niet te bepalen gemiddelde omvang (bij 1, 2 en 4 kHz)'
    elif LG_1_4 < 20 and LG_05_1 < 20: type_hearing_loss = '(nagenoeg) normale gehoordrempels, ' + str(LG_1_4) + after_extent_hl
    elif LG_1_4 < 20 and 20 <= LG_05_1 < 25: type_hearing_loss = 'nagenoeg normale gehoordrempels, ' + str(LG_1_4) + after_extent_hl
    elif LG_1_4 >= 20 and ABG == None: type_hearing_loss = '(vermoedelijk) perceptief / conductief / gemengd gehoorverlies, ' + str(LG_1_4) + after_extent_hl
    elif LG_1_4 >= 20 and ABG < 15: type_hearing_loss = 'perceptief gehoorverlies, ' + str(LG_1_4) + after_extent_hl
    elif LG_1_4 >= 20 and ABG > 15 and BG_05_1 <= 15: type_hearing_loss = 'conductief gehoorverlies, ' + str(LG_1_4) + after_extent_hl
    elif LG_1_4 >= 20 and ABG > 15 and 15 < BG_05_1 <= 20: type_hearing_loss = '(overwegend) conductief gehoorverlies, ' + str(LG_1_4) + after_extent_hl
    elif LG_1_4 >= 20 and ABG > 15 and BG_05_1 > 20: type_hearing_loss = 'gemengd gehoorverlies, ' + str(LG_1_4) + after_extent_hl
    
    return type_hearing_loss


# In[14]:


def short_hearing_loss(hearing_loss_data):
    """
    Generates a short description of the hearing loss based on provided data.

    Args:
        hearing_loss_data (tuple): Tuple containing hearing loss data (BG_05_1, LG_05_1, LG_1_4, ABG).

    Returns:
        str: Short description of hearing loss extent and type.
    """
    BG_05_1, LG_05_1, LG_1_4, ABG = hearing_loss_data
    
    if LG_1_4 == None: type_hearing_loss = ' gemengd/perceptief gehoorverlies / (functioneel) doof oor'
    elif LG_1_4 < 20 and LG_05_1 < 20: type_hearing_loss = '(nagenoeg) normale gehoordrempels'
    elif LG_1_4 >= 20 and ABG == None: type_hearing_loss = ' (vermoedelijk) perceptief / conductief / gemengd gehoorverlies'
    elif LG_1_4 >= 20 and ABG < 15: type_hearing_loss = ' perceptief gehoorverlies'
    elif LG_1_4 >= 20 and ABG > 15 and BG_05_1 <= 15: type_hearing_loss = ' conductief gehoorverlies'
    elif LG_1_4 >= 20 and ABG > 15 and 15 < BG_05_1 <= 20: type_hearing_loss = ' (overwegend) conductief gehoorverlies'
    elif LG_1_4 >= 20 and ABG > 15 and BG_05_1 > 20: type_hearing_loss = ' gemengd gehoorverlies'
 

    if LG_1_4 == None: 
        extent_hl = 'een zeer ernstig'
    elif LG_1_4 < LIGHT_EXTENT: 
        extent_hl = ''
    elif LIGHT_EXTENT <= LG_1_4 < LIGHT_TO_MODERATE_EXTENT: 
        extent_hl = 'een licht'
    elif LIGHT_TO_MODERATE_EXTENT <= LG_1_4 < MODERATE_EXTENT: 
        extent_hl = 'een licht tot matig'
    elif MODERATE_EXTENT <= LG_1_4 < MODERATE_TO_SEVERE_EXTENT: 
        extent_hl = 'een matig'
    elif MODERATE_TO_SEVERE_EXTENT <= LG_1_4 < SEVERE_EXTENT: 
        extent_hl = 'een matig tot ernstig'
    elif SEVERE_EXTENT <= LG_1_4 < SEVERE_TO_VERY_SEVERE_EXTENT: 
        extent_hl = 'een ernstig'
    elif SEVERE_TO_VERY_SEVERE_EXTENT <= LG_1_4 < VERY_SEVERE_EXTENT: 
        extent_hl = 'een ernstig tot zeer ernstig'
    elif LG_1_4 >= VERY_SEVERE_EXTENT: 
        extent_hl = 'een zeer ernstig'
  
    
    return (extent_hl + type_hearing_loss)


# In[15]:


def patient_data(pdf_infile):
    """
    Extracts patient data (pt-ID and birthday) from the header of a PDF audiogram document.

    Args:
        pdf_infile (str): Path to the PDF file.

    Returns:
        str: Formatted patient data containing pt-ID and birthday.
    """
    ...
    header = extract_header_text(pdf_infile)
    l = header.split('\n')
        # Define the patterns to search for
    patterns = ['PatientId:', 'Geb-dat:']

    # Initialize the variables with None
    pt_id = None
    birthday = None

    # Traverse list l to find hearing loss data
    for element in l:
        for pattern in patterns:
            if pattern in element:
                try:
                    value = element.split(pattern)[1]  # Take the string
                except IndexError:
                    value = None  # Set to None when extraction fails

                if pattern == 'PatientId:':
                    pt_id = value
                elif pattern == 'Geb-dat:':
                    birthday = value
                break  # Break the loop after finding the first occurrence
    
    result = 'PtId_' + pt_id + '_GebDat_' + birthday
    result = result.replace(' ', '')
    return result


# In[16]:


def test_date(pdf_infile):
    """
    Extracts the test date from the bottom text of a PDF audiogram document.

    Args:
        pdf_infile (str): Path to the PDF file.

    Returns:
        str or None: Extracted test date in 'dd-mm-yyyy' format or None if not found.
    """
    text = extract_bottom_text(pdf_infile)
    # Define a regular expression pattern to match dates in the format 'dd-mm-yyyy'
    date_pattern = r'\d{2}-\d{2}-\d{4}'
    
    # Find all matches of the pattern in the text
    dates = re.findall(date_pattern, text)
    
    # Return the first match if found, otherwise return None
    if dates:
        return dates[0]
    else:
        return None   


# In[17]:


def print_date(pdf_infile):
    """
    Extracts the print date from the bottom text of a PDF audiogram document.

    Args:
        pdf_infile (str): Path to the PDF file.

    Returns:
        str or None: Extracted test date in 'dd-mm-yyyy' format or None if not found.
    """
    text = extract_bottom_text(pdf_infile)
    # Define a regular expression pattern to match dates in the format 'dd-mm-yyyy'
    date_pattern = r'\d{2}-\d{2}-\d{4}'
    
    # Find all matches of the pattern in the text
    dates = re.findall(date_pattern, text)
    
    # Return the first match if found, otherwise return None
    if dates:
        return dates[1]
    else:
        return None   


# In[18]:


def check_for_insert(pdf_infile):
    """
    Checks for the presence of string 'insert' in the bottom text of a PDF audiogram document.

    Args:
        pdf_infile (str): Path to the PDF file.

    Returns:
        str: Returns 'met insert phones ' if 'insert' is found, else an empty string.
    """
    text = extract_bottom_text(pdf_infile)
    text = text.lower()

    res = None

    if 'insert' in text: res = 'met insert phones'
    else: res = ''  

    return res


# In[19]:


def check_for_hoofdtel(pdf_infile):
    """
    Checks for the presence of string 'hoofdtel' or 'koptel' in the bottom text of a PDF audiogram document.

    Args:
        pdf_infile (str): Path to the PDF file.

    Returns:
        str: Returns 'met hoofdtelefoon' if 'koptel' is found, else an empty string.
    """
    text = extract_bottom_text(pdf_infile)
    text = text.lower()

    res = None

    if 'hoofdtel' in text or 'koptel' in text: res = 'met hoofdtelefoon'
    else: res = ''  

    return res


# In[20]:


def check_for_bloktest(pdf_infile):
    """
    Checks for the presence of string 'blok' or 'bt' in the bottom text of a PDF audiogram document.

    Args:
        pdf_infile (str): Path to the PDF file.

    Returns:
        str: Returns 'blokkentest ' if 'blok' (or 'bt') is found, else an empty string.
    """
    text = extract_bottom_text(pdf_infile)
    text = text.lower()

    res = None

    if 'blok' in text: res = 'blokkentest'
    elif 'bt' in text: res = 'blokkentest'
    else: res = ''  

    return res


# In[21]:


def sum_of_checks(pdf_infile):
    ch = check_for_hoofdtel(pdf_infile)
    ci = check_for_insert(pdf_infile)
    cb = check_for_bloktest(pdf_infile)
    t = cb, ci, ch
    l = [s for s in t if s != '']
    res = ' '.join(l)
    return res


# In[22]:


def check_for_speech_material(pdf_infile):
    """
    Checks for the presence of string 'blok' or 'bt' in the bottom text of a PDF audiogram document.

    Args:
        pdf_infile (str): Path to the PDF file.

    Returns:
        str: Returns 'blokkentest ' if 'blok' (or 'bt') is found, else an empty string.
    """
    text = extract_bottom_text(pdf_infile)
    text = text.lower()

    res = None

    if 'sap' in text: res = 'SAP '
    elif 'pit' in text: res = 'PIT'
    elif 'kl' in text or 'kind' in text: res = 'NVA-kinderlijsten'
    else: res = 'NVA'  

    return res


# In[23]:


def patient_sex(pdf_infile):
    """
    Extracts patient sex information from the header of a PDF audiogram document.

    Args:
        pdf_infile (str): Path to the PDF file.

    Returns:
        str: Returns 'Mevrouw ' if patient is female, 'De heer ' if patient is male.
    """
    header = extract_header_text(pdf_infile)
    l = header.split('\n')
        # Define the patterns to search for
    patterns = ['Vrouw', 'Man']

    # Initialize the variable with None
    sex = None

    # Traverse list l to find hearing loss data
    for element in l:
        for pattern in patterns:
            if pattern in element:
                try:
                    value = pattern  # Take the string
                except IndexError:
                    value = None  # Set to None when extraction fails

                if value == 'Vrouw':
                    sex = 'Mevrouw '
                elif value == 'Man':
                    sex = 'De heer '
                break  # Break the loop after finding the first occurrence
    
    return sex


# In[24]:


def extract_age_from_header(pdf_infile):
    """
    Extracts the age (in years) from the header text.

    Args:
        pdf_infile (str): Path to the PDF file containing the header text.

    Returns:
        int or None: Extracted age in years or None if not found.
    """
    header_text = extract_header_text(pdf_infile)
    
    # Define the pattern to search for
    age_pattern = r'Leeftijd: (\d+)'

    # Search for the age pattern in the header text
    match = re.search(age_pattern, header_text)

    if match:
        # Extract the matched age
        age_in_years = int(match.group(1))
        return age_in_years

    return None  # Return None if the age pattern is not found


# In[25]:


def extract_surname_from_pdf(pdf_infile):
    """
    Extracts the surname from the header text of a PDF document.

    Args:
        pdf_infile (str): Path to the PDF file.

    Returns:
        str: Extracted surname.
    """
    # Extract header text from the PDF
    header_text = extract_header_text(pdf_infile)

    # Split the header text into lines
    lines = header_text.split('\n')
    line = lines[2] # Search for the line containing the surname and family name
    parts = line.split(',') # Split the line into parts using the comma as a separator
        
    # Assuming the surname is the second part (index 1) after the comma
    if len(parts) > 1:
        surname = parts[1].strip()
        return surname + ' '

    # If the surname is not found, return an empty string
    return ""


# In[26]:


def crop_pdf(open_pdf_document, crop_rect):
    """
    Crop a PDF file and return the cropped pdf_document.

    Args:
        pdf_document (fitz.Document): PyMuPDF document object.
        crop_rect (tuple): Coordinates (left, top, right, bottom) to crop the PDF.

    Returns:
        fitz.Document or None: Cropped PyMuPDF document object if successful, else None.
    """
    try:
        
        # Crop the PDF using the specified rectangle
        open_pdf_document[0].set_cropbox(fitz.Rect(crop_rect))

        # print("PDF successfully cropped")
        return open_pdf_document
    
    except Exception as e:
        print(f"An error occurred while cropping the PDF: {e}")
        return None


# In[27]:


def filter_colored_points(image, color_range):
    """
    Filter colored points in an image based on the specified RGB range.

    Args:
        image (PIL.Image.Image): The input image.
        color_range (list of tuples): RGB ranges for filtering colored points.

    Returns:
        list of tuples: Coordinates of the selected colored points.
    """
    # Convert the image to a NumPy array
    img_array = np.array(image)

    # Extract the RGB channels
    r_channel = img_array[:, :, 0]
    g_channel = img_array[:, :, 1]
    b_channel = img_array[:, :, 2]

    # Create masks for each channel based on the specified ranges
    color_mask = (
        (r_channel >= color_range[0][0]) & (r_channel <= color_range[0][1]) &
        (g_channel >= color_range[1][0]) & (g_channel <= color_range[1][1]) &
        (b_channel >= color_range[2][0]) & (b_channel <= color_range[2][1])
    )

    # Get the coordinates of the selected points
    colored_points_coords = np.argwhere(color_mask)

    return colored_points_coords


# In[28]:


def filter_red_points(image):
    """
    Filter red points in an image based on a predefined RGB range.

    Args:
        image (PIL.Image.Image): The input image.

    Returns:
        list of tuples: Coordinates of the selected red points.
    """
    result = filter_colored_points(image, RED_COLOR_RANGE)
    return result 

def red_points_coords(pdf_path, crop_rect = (SPEECH_RE_LEFT, SPEECH_RE_TOP, SPEECH_RE_RIGHT, SPEECH_RE_BOTTOM)):
    
    # Crop the PDF to get the speech audiogram for the right ear
    speech_audio_right = crop_pdf(pdf_path, crop_rect)

    if speech_audio_right is not None:
        try:
            # Render the page as an image
            pix = speech_audio_right[0].get_pixmap()
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

            # Get the coordinates of red points
            red_points_coords = filter_red_points(img)

            # Count the number of red points
            # red_points_count = len(red_points_coords)

            # print(f"Number of Red Points: {red_points_count}")

        except Exception as e:
            print(f"Error processing PDF: {e}")

    return red_points_coords


# In[29]:


def filter_blue_points(image):
    """
    Filter blue points in an image based on a pblueefined RGB range.

    Args:
        image (PIL.Image.Image): The input image.

    Returns:
        list of tuples: Coordinates of the selected blue points.
    """
    result = filter_colored_points(image, BLUE_COLOR_RANGE)
    return result 

def blue_points_coords(pdf_path, crop_rect = (SPEECH_LE_LEFT, SPEECH_LE_TOP, SPEECH_LE_RIGHT, SPEECH_LE_BOTTOM)):
    
    # Crop the PDF to get the speech audiogram for the right ear
    speech_audio_right = crop_pdf(pdf_path, crop_rect)

    if speech_audio_right is not None:
        try:
            # Render the page as an image
            pix = speech_audio_right[0].get_pixmap()
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

            # Get the coordinates of blue points
            blue_points_coords = filter_blue_points(img)

            # Count the number of blue points
            # blue_points_count = len(blue_points_coords)

            # print(f"Number of blue Points: {blue_points_count}")

        except Exception as e:
            print(f"Error processing PDF: {e}")

    return blue_points_coords


# In[30]:


def transform_coordinates_right_speechaudio(coords_array, origin_point = SPEECH_RE_ORIGIN):
    """
    Transform coordinates from the right speech audiogram to a standardized format.

    Parameters:
    - coords_array (numpy.ndarray): Input array with shape (n, 2) representing (x, y) coordinates.
    - origin_point (tuple): Origin point for coordinate transformation.

    Returns:
    numpy.ndarray: Transformed coordinates in a standardized format.

    This function adjusts the x and y coordinates based on the provided origin_point.
    It scales the x coordinate to represent dB levels and the y coordinate to represent speech scores.
    The transformed coordinates are returned as a numpy array with shape (n, 2).

    Example:
    >>> origin = (44, 118)
    >>> coordinates = np.array([[50, 90], [60, 80], [70, 70]])
    >>> transform_coordinates_right_speechaudio(coordinates, origin)
    array([[ 47.83,   0.  ],
           [ 57.5 ,   0.22],
           [ 67.17,   0.44]])
    """
    # Adjust x coordinates for origin, swapping them too
    dB_levels = 3 + (coords_array[:, 1] - origin_point[0]) * 120 / (LEVEL_120_RE - LEVEL_0_RE)

    # Adjust y coordinates for origin, swapping them too
    score = (origin_point[1] - coords_array[:, 0]) / (SCORE_0 - SCORE_100)

    result = np.column_stack((dB_levels, score))  # Create a numpy array [x, y]
    return result


# In[31]:


def transform_coordinates_left_speechaudio(coords_array, origin_point = SPEECH_LE_ORIGIN):
    """
    Transform coordinates from the left speech audiogram to a standardized format.

    Parameters:
    - coords_array (numpy.ndarray): Input array with shape (n, 2) representing (x, y) coordinates.
    - origin_point (tuple): Origin point for coordinate transformation.

    Returns:
    numpy.ndarray: Transformed coordinates in a standardized format.

    This function adjusts the x and y coordinates based on the provided origin_point.
    It scales the x coordinate to represent dB levels and the y coordinate to represent speech scores.
    The transformed coordinates are returned as a numpy array with shape (n, 2).

    Example:
    >>> origin = (26, 118)
    >>> coordinates = np.array([[50, 90], [60, 80], [70, 70]])
    >>> transform_coordinates_right_speechaudio(coordinates, origin)
    array([[ 47.83,   0.  ],
           [ 57.5 ,   0.22],
           [ 67.17,   0.44]])
    """
    # Adjust x coordinates for origin, swapping them too
    dB_levels = 3 + (coords_array[:, 1] - origin_point[0]) * 120 / (LEVEL_120_LE - LEVEL_0_LE)

    # Adjust y coordinates for origin, swapping them too
    score = (origin_point[1] - coords_array[:, 0]) / (SCORE_0 - SCORE_100)

    result = np.column_stack((dB_levels, score))  # Create a numpy array [x, y]
    return result


# In[32]:


def find_max_score_and_level(arr_adj_coords):
    """
    Analyze a 2D numpy array with speech levels and speech scores.
    
    Returns the maximum speech score and its corresponding level.

    Parameters:
    - arr_adj_coords (numpy.ndarray): 2D array with shape (n, 2) representing (speech level, speech score).

    Returns:
    tuple: A tuple containing the speech level and the corresponding maximal speech score in %.

    If the array is empty, returns (None, None).
    """
    if arr_adj_coords.size == 0:
        return (None, None)  # Handle the case with an empty array

    # Find the index of the element containing the max speech score
    id_max = arr_adj_coords[:, 1].argmax()
    
    # Extract the array element with max speech score
    arr_max = arr_adj_coords[id_max]

    # Round and adjust the max score
    score_max = arr_max[1].round(2)
    score_max = int(1 + ((score_max * 100 / 3).round(0)) * 3) if score_max > 0.49 else int(((score_max * 100 / 3).round(0)) * 3)
    score_max = 100 if score_max > 100 else score_max
    score_max = 0 if score_max < 0 else score_max
    
    # Round and adjust the max level
    level_max = int(((arr_max[0] / 5).round(0)) * 5)
    level_max = 120 if level_max > 120 else level_max
    level_max = 0 if level_max < 0 else level_max
    
    return level_max, score_max


# In[33]:


def find_srt_level(arr_adj_coords):
    """
    Find the Speech Reception Threshold (SRT) level corresponding to a speech score of 0.5.

    Parameters:
    - arr_adj_coords (numpy.ndarray): Input array with shape (n, 2) representing (x, y) coordinates.

    Returns:
    float or None: SRT level if it can be determined, else None.

    This function performs linear interpolation to estimate the SRT level based on the input coordinates.
    If the maximal speech score is less than 0.5 or the minimal speech score is greater than 0.5,
    the function returns None.
    """
    if arr_adj_coords.size == 0:
        return None  # Handle the case with an empty array

    # Sort the array by speech score
    sorted_coords = arr_adj_coords[arr_adj_coords[:, 1].argsort()]

    # Extract the dB levels and speech scores
    dB_levels = sorted_coords[:, 0]
    scores = sorted_coords[:, 1]

    # Check if the target score (0.5) is within the range
    if scores[0] > 0.5 or scores[-1] < 0.5:
        return None

    # Find the index where the transition occurs
    transition_index = np.searchsorted(scores, 0.5, side="right")

    # Perform linear interpolation
    srt_level_abs = np.interp(0.5, scores[transition_index - 1:transition_index + 1], dB_levels[transition_index - 1:transition_index + 1])
    srt_level = int(srt_level_abs - 25)
    
    return srt_level


# In[34]:


def describe_speech_audiogram(level_max_score, max_score, srt, side):
    """
    Generate a descriptive text for a speech audiogram based on three features:
    
    Parameters:
    - level_max_score (float or None): Speech level corresponding to the maximal speech score.
    - max_score (float): Maximal speech score in the speech audiogram.
    - srt (float or None): Speech Reception Threshold level corresponding to a speech score of 0.5.

    Returns:
    str: A descriptive text based on the provided features.
    If level_max_score is None, returns 'niet getest of spraakverstaan nihil'.
    If max_score is less than 50, adds ', drempelverschuiving niet te bepalen' to the text.
    If srt is None, adds ', drempelverschuiving niet bepaald' to the text.
    Otherwise, adds ', drempelverschuiving {srt} dB' to the text.

    Example:
    >>> describe_speech_audiogram(60, 75, 25)
    'maximale score 75% bij 60 dB, drempelverschuiving 25 dB'
    """
    if level_max_score is None:
        return side + ' ' +  'niet getest of spraakverstaan nihil'

    text_srt = ', drempelverschuiving niet bepaald' if srt is None else f', drempelverschuiving {srt} dB'
    text_score = f'maximale score {max_score}% bij {level_max_score} dB'
    
    if max_score < 50:
        return side + ' ' + text_score + ', drempelverschuiving niet te bepalen'
    
    return side + ' ' +  text_score + text_srt


# In[35]:


def toonaudiometry_results(pdf_infile, hearing_loss_data_right_ear, hearing_loss_data_left_ear, speech_audio_result):
    """
    Generates a formatted report of the audiometry results.

    Args:
        pdf_infile (str): Path to the PDF file.

    Returns:
        str: Formatted report of the audiometry results.
    """
    try:
        # data of test and patient
        pt = '\n' + patient_data(pdf_infile)
        day = '\nOnderzoeksdatum: ' + test_date(pdf_infile) + ', Afdrukdatum: ' + print_date(pdf_infile) + '\n'
        date = '\nDatum: ' + test_date(pdf_infile)
        
        # text with comments in the bottom section
        text = "\n\nOpmerkingen\n" + extract_bottom_text(pdf_infile)
        
        # check for age of subject
        age_at_test = extract_age_from_header(pdf_infile)
        if age_at_test > 18: nm = patient_sex(pdf_infile)
        else: nm = extract_surname_from_pdf(pdf_infile)
        
        # compose lines of text for "long description"
        a = "\nToonaudiometrie " + '(' + sum_of_checks(pdf_infile)+ ')'
        b = "\nBetrouwbaarheid goed/redelijk/matig/slecht"
        c = "\nRechts " + type_hearing_loss(hearing_loss_data_right_ear)
        d = "\nLinks " + type_hearing_loss(hearing_loss_data_left_ear) + "\n\n"
        g = f"Spraakaudiometrie ({check_for_speech_material(pdf_infile)}{' ' if check_for_hoofdtel(pdf_infile) != '' or check_for_insert(pdf_infile) != '' else ''}{check_for_hoofdtel(pdf_infile)}{check_for_insert(pdf_infile)})\n"
        h = speech_audio_result + "\n\n"
        
        # compose lines of text for "short description"
        z = '\n' + patient_sex(pdf_infile) + 'heeft rechts ' + short_hearing_loss(hearing_loss_data_right_ear) \
        + ' en links ' + short_hearing_loss(hearing_loss_data_left_ear) + '.'
        if short_hearing_loss(hearing_loss_data_right_ear) == short_hearing_loss(hearing_loss_data_left_ear):
            z = '\n' + nm + 'heeft beiderzijds ' + short_hearing_loss(hearing_loss_data_right_ear) + '.\n'
        else: 
            z = '\n' + nm + 'heeft rechts ' + short_hearing_loss(hearing_loss_data_right_ear) \
            + ' en links ' + short_hearing_loss(hearing_loss_data_left_ear) + '.\n'
            
        # compose the all text
        result = pt + day + date + a + b + c + d + g + h + date + z + text
        print(result)
        return result

    except Exception as e:
        print(f"An error occurred while generating audiometry results: {e}")
        return ""


# In[36]:


def get_pdf_file_path_from_user():
    """
    Prompts the user to enter the path to the PDF file containing the audiogram.

    Returns:
        str: The user-provided path to the PDF file.
    """
    while True:
        pdf_path = input(Fore.LIGHTRED_EX + Style.BRIGHT +"\nVoer de locatie van het PDF-bestand met het audiogram in:\n"\
                         + Style.RESET_ALL)

        # Check if the path exists
        if not os.path.exists(pdf_path):
            print("Ongeldige locatie. Het opgegeven pad bestaat niet. Probeer opnieuw.")
            continue

        # Check if the path is a file
        if not os.path.isfile(pdf_path):
            print("Ongeldige locatie. Het opgegeven pad leidt niet naar een bestand. Probeer opnieuw.")
            continue

        # Check if the file has a PDF extension
        if not pdf_path.lower().endswith('.pdf'):
            print("Ongeldig bestandsformaat. Het bestand moet een PDF-bestand zijn. Probeer opnieuw.")
            continue

        return pdf_path


# In[37]:


def get_user_downloads_folder():
    """
    Get the appropriate Downloads folder based on the user's setup.
    
    Returns:
        str: Path to the user's Downloads folder.
    """
    # Get the path to the user's standard Downloads folder
    standard_downloads = join(expanduser("~"), "Downloads")
    
    # Check if the standard Downloads folder exists
    if os.path.exists(standard_downloads):
        return standard_downloads
    
    # If standard Downloads folder doesn't exist, try the OneDrive Downloads folder
    onedrive_downloads = join(expanduser("~"), "OneDrive - adelante-zorggroep.nl", "Downloads")
    
    # Check if the OneDrive Downloads folder exists
    if os.path.exists(onedrive_downloads):
        return onedrive_downloads
    
    # If neither exists, return None or raise an exception based on your preference
    raise FileNotFoundError("Neither standard Downloads nor OneDrive Downloads folder found.")

def save_results_to_downloads(results, pdf_infile):
    """
    Saves the formatted results to a text file in the user's preferred Downloads folder.

    Args:
        results (str): Formatted results.
        pdf_infile (str): Path to the PDF file.
    """
    try:
        # Get the user's preferred Downloads folder
        downloads_folder = get_user_downloads_folder()
        
        # Define the file name
        f_name = patient_data(pdf_infile) + '_' + 'TestDat_' + test_date(pdf_infile) + '.txt'
        
        # Create the full path to the file in the user's preferred Downloads folder
        file_path = join(downloads_folder, f_name)

        # Save the formatted results to the output file
        with open(file_path, 'w') as file:
            file.write(results)

        print("\nResultaten opgeslagen in: \n", file_path)

    except Exception as e:
        print(f"An error occurred while saving results to file: {e}")


# In[38]:


# Main program
def main(): 

    while True:
        # Ask the user to provide the location of the PDF file with an audiogram to analyze
        pdf_infile = get_pdf_file_path_from_user()

        # Open PDF safely
        pdf_document = open_pdf_safely(pdf_infile)

        if pdf_document is not None:
            try:
                # Extract text from right_ear and left_ear audiograms
                right_ear_audiogram_text = extract_right_audiogram_text(pdf_document)
                left_ear_audiogram_text = extract_left_audiogram_text(pdf_document)

                # Extract and Process Hearing Loss data
                hearing_loss_data_right_ear = provide_hearing_loss_data(right_ear_audiogram_text)
                hearing_loss_data_left_ear = provide_hearing_loss_data(left_ear_audiogram_text)

                # Extract and Process Right ear Speech-audiogram
                red_points_re = red_points_coords(pdf_document, crop_rect=(SPEECH_RE_LEFT, SPEECH_RE_TOP, SPEECH_RE_RIGHT, SPEECH_RE_BOTTOM))
                adj_red_points_re = transform_coordinates_right_speechaudio(red_points_re)
                level_max_score_re, max_score_re = find_max_score_and_level(adj_red_points_re)
                srt_re = find_srt_level(adj_red_points_re)
                text_re = describe_speech_audiogram(level_max_score_re, max_score_re, srt_re, side='Rechts' )

                # Extract and Process Left ear Speech-audiogram
                blue_points_le = blue_points_coords(pdf_document, crop_rect=(SPEECH_LE_LEFT, SPEECH_LE_TOP, SPEECH_LE_RIGHT, SPEECH_LE_BOTTOM))
                adj_blue_points_le = transform_coordinates_left_speechaudio(blue_points_le)
                level_max_score_le, max_score_le = find_max_score_and_level(adj_blue_points_le)
                srt_le = find_srt_level(adj_blue_points_le)
                text_le = describe_speech_audiogram(level_max_score_le, max_score_le, srt_le, side='Links')

                # Speech-audiometry Result
                speech_audiometry_text = f'{text_re}\n{text_le}'

                # Get the formatted result
                formatted_result = toonaudiometry_results(pdf_document, hearing_loss_data_right_ear, hearing_loss_data_left_ear, speech_audiometry_text)

                # Save the formatted result to a text file in the same directory as the PDF file
                save_results_to_downloads(formatted_result, pdf_infile)
                
                
            except Exception as e:
                print(f"An error occurred: {e}")

        # Check if the user wants to exit
        exit_input = input(Fore.LIGHTRED_EX + Style.BRIGHT +"\nDruk op Enter voor de analyse van een andere file, of typ 'exit' en druk op Enter om te verlaten:\n"\
                           + Style.RESET_ALL)
        if exit_input.lower() == 'exit':
            print("Het programma wordt afgesloten. Tot ziens!")
            break


# In[39]:


#   /home/guido/Python_Programs/project_audiogram/EasyPDF_AudioScraper/bad74547-e6be-41c0-8f8d-bdac23416730.pdf
# Call the main function to execute the program
if __name__ == "__main__":
    main()


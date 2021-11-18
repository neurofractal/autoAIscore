#!/usr/bin/env python3

# Import packages
from docx import Document
from lxml import etree
import zipfile
import re
import os
import sys
import pandas as pd
import warnings

def custom_formatwarning(msg, *args, **kwargs):
    # ignore everything except the message
    return str(msg) + '\n'

warnings.formatwarning = custom_formatwarning

ooXMLns = {'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
#Function to extract all the comments of document(Same as accepted answer)
#Returns a dictionary with comment id as key and comment string as value
def get_document_comments(docxFileName):
    comments_dict={}
    docxZip = zipfile.ZipFile(docxFileName)
    commentsXML = docxZip.read('word/comments.xml')
    et = etree.XML(commentsXML)
    comments = et.xpath('//w:comment',namespaces=ooXMLns)
    for c in comments:
        comment=c.xpath('string(.)',namespaces=ooXMLns)
        comment_id=c.xpath('@w:id',namespaces=ooXMLns)[0]
        comments_dict[comment_id]=comment
    return comments_dict

def paragraph_comments(paragraph,comments_dict):
    comments=[]
    for run in paragraph.runs:
        comment_reference=run._r.xpath("./w:commentReference")
        if comment_reference:
            comment_id=comment_reference[0].xpath('@w:id',namespaces=ooXMLns)[0]
            comment=comments_dict[comment_id]
            comments.append(comment)
    return comments


def extract_AI_scores(docxFileName):

    document = Document(docxFileName)
    comments_dict=get_document_comments(docxFileName)
    
    feature_text = []
    category = []
    sub_category = []
    accuracy = []
    event_number = []
    
    event_number_within_loop = []
    
    for para in document.paragraphs:
        # Check if this paragraph contains Event X, if so update event_number_within_loop
        find_event_number = (re.findall(r'\bEvent \d+',para.text))
        if find_event_number:
            event_number_within_loop = [int(s) for s in find_event_number[0].split() if s.isdigit()][0]
            print("Processing Event {}".format(event_number_within_loop))
        
        # Check if paragraph contains a comment
        comm = []
        for run in para.runs:
            comment_reference=run._r.xpath("./w:commentReference")
            if comment_reference:
                comm = comment_reference
        
        # If it does...
        if comm:
    #       Add paragraph text   
            feature_text.append(para.text)

    #       Warn the user if there is an unusually short string
            if len(para.text) < 1:
                warnings.warn("TEXT TOO SHORT. Event: {} Detail: '{}'".format(event_number_within_loop,para.text))

    #       Get comment from this paragraph
            r = paragraph_comments(para,comments_dict)
    #       Add the text to feature_text

    #       Search for pattern of characters

            patt = re.findall(r'[IEie][ETPRSOetprso][VILEHTvileht][TFUtfu]',r[0])

    #       If pattern is found...
            if patt:
                # Warn if length is not 4
                if len(patt[0]) != 4:
                    warnings.warn("WEIRD LENGTH. Event: {} Detail: '{}'".format(event_number_within_loop,para.text))
                
                # Category: I = internal ; E = external
                if patt[0][0].upper() == 'I':
                    text1 = 'internal'
                elif patt[0][0].upper() == 'E':
                    text1 = 'external'
                
                category.append(text1)      
                
                # Sub-Category: EV = event ; PE = perceptual ; TI = time ; PL = Place
                #               TH = thought_emotion ; SE = semantic ; RE = repetition ;
                #               OT = other
                                
                if patt[0][1:3].upper() == 'EV':
                    text2 = 'event'
                elif patt[0][1:3].upper() == 'PE':
                    text2 = 'perceptual'
                elif patt[0][1:3].upper() == 'TI':
                    text2 = 'time'
                elif patt[0][1:3].upper() == 'PL':
                    text2 = 'place'
                elif patt[0][1:3].upper() == 'TH':
                    text2 = 'thought_emotion'
                elif patt[0][1:3].upper() == 'SE':
                    text2 = 'semantic'
                elif patt[0][1:3].upper() == 'RE':
                    text2 = 'repetition'
                elif patt[0][1:3].upper() == 'OT':
                    text2 = 'other'
                else:
                    warnings.warn("WEIRD SUB-CATEGORY PATTERN FOUND. Event: {} Detail: '{}'".format(event_number_within_loop,para.text))
                    text2 = ''
                
                sub_category.append(text2)
                
                # Accuracy: T = true ; F = false ; U = unverifiable
                if patt[0][3].upper() == 'T':
                    text3 = 'true'
                elif patt[0][3].upper() == 'F':
                    text3 = 'false'
                elif patt[0][3].upper() == 'U':
                    text3 = 'unverifiable'
                    
                accuracy.append(text3)
                
                # Add event number
                event_number.append(event_number_within_loop)
                
    #       If pattern is not found make NaNs and create warning...
            else:    
                category.append('NaN')       
                sub_category.append('NaN')
                accuracy.append('NaN')
                event_number.append(event_number_within_loop)

                warnings.warn("PATTERN NOT FOUND. Event: {} Detail: '{}'".format(event_number_within_loop,para.text))
        
        
    # Create data frame
    df = pd.DataFrame(list(zip(event_number,category,sub_category,accuracy,feature_text)),
                   columns =['event_number','category','sub_category','accuracy','text'])

    return df
 

# Get input and output document
inDoc = sys.argv[1]
print("")
print("Processing File: {}".format(inDoc))
print("")
filename, file_extension = os.path.splitext(sys.argv[1])
outFolder = "{}.csv".format(filename)

df = extract_AI_scores(inDoc)


# Print the head
df.to_csv(outFolder)
print("")
print("Output: {}".format(outFolder))
print("")


     
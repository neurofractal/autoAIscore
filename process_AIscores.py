#!/usr/bin/env python3

# Import packages
from docx import Document
from lxml import etree
import zipfile
import re
import glob
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

    event_number = (re.findall(r'\d{2}',docxFileName))
    event_number = int(event_number[0])
    
    feature_text = []
    category = []
    sub_category = []
    accuracy = []

    for para in document.paragraphs:
        comm = []
        for run in para.runs:
            comment_reference=run._r.xpath("./w:commentReference")
            if comment_reference:
                comm = comment_reference
        if comm:
    #       Add paragraph text   
            feature_text.append(para.text)
                    
    #       Warn the user if there is an unusually short string
            if len(para.text) < 1:
                warnings.warn("   TEXT TOO SHORT. Event: {} Detail: '{}'".format(event_number,para.text))

    #       Get comment from this paragraph
            r = paragraph_comments(para,comments_dict)
    #       Add the text to feature_text

    #       Search for pattern of character, number, character
    #		This could be improved by just searching for characters matching the AI scoring code
            patt = re.findall(r'\w\d\w+',r[0])

    #       If pattern is found...
            if patt:
                # Warn if length of this is not 3
                if len(patt[0]) != 3:
                    warnings.warn("WEIRD LENGTH. Event: {} Detail: '{}'".format(event_number,para.text))

                # Add to appropriate []
                category.append(patt[0][0].capitalize())       
                sub_category.append(int(patt[0][1]))
                accuracy.append(patt[0][2].capitalize())

    #       If pattern is not found make NaNs and create warning...
            else:    
                category.append('NaN')       
                sub_category.append('NaN')
                accuracy.append('NaN')

                warnings.warn("PATTERN NOT FOUND. Event: {} Detail: '{}'".format(event_number,para.text))

    # Create data frame
    df = pd.DataFrame(list(zip([event_number]*len(category),category,sub_category,accuracy,feature_text)),
                   columns =['event_number','category','sub_category','accuracy','text'])
        
    return df
 

# Get input and output folder
inFolder = sys.argv[1]
print("")
print("Processing Folder: {}".format(inFolder))
print("")
outFolder = sys.argv[2]

#   
df = pd.DataFrame(columns=['event_number','category', 
                            'sub_category','accuracy',
                            'text'])

for file in sorted(glob.glob("{}{}".format(inFolder,'/*.docx'))):
    # print(file)
    df_event = extract_AI_scores(file)
    df = pd.concat([df,df_event],sort=False)



# Print the head
df.head()
df.to_csv(outFolder)
print("")
print("Output: {}".format(outFolder))
print("")


     
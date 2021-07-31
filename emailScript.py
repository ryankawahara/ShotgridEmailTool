import os
import shotgun_api3
import json

import sys
import urllib
import logging
import urllib.parse
import smtplib
import string
import array as arr
import pandas as pd
from collections import defaultdict
# https://www.geeksforgeeks.org/python-convert-list-of-dictionaries-to-dictionary-value-list/

from email import generator
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from email import encoders
from email.mime.base import MIMEBase


from datetime import datetime
from datetime import date

import re
import subprocess
import io
import glob

from emailSetup import *

#this path is where the .eml template and log will be sent
desktopPath = os.path.dirname(__file__)
filePath = desktopPath + "/submissionDocs/"
projectAttributes = []
attributesToReceive = []
toggleSubmission=True
incrementListIndex=False
arrayOfLineswithAttributes = []
paragraphAttributes=[]
paragraphAttributeArray=[]
paragraphAttributeArrayFull=[]
startingValue=0
multiLineEntry=False
attributeLength=0





def _html_header():
    return """
<html>
<head>
<style>
div.gallery {
}

div.desc{
}
</style>
<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.0.1/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-+0n0xVW2eSR5OomGNYDnhzAbDsOXxcvSN1TPprVMTNDbiYZCxYbOOl7+AMvyTG2x" crossorigin="anonymous">
</head>
<body style="margin-left:35px">
"""

def greetingLine(projectName):
    return """
     %s <br>
    """ % (projectName)

def _image_html(link, image):
    return """
<br>
<div class = "gallery">
    <a href = "%s">
        <img src="%s">
    </a>
</div>

    """ % (link, image)

def _shot_name_html(title):
    return """
    <div class = "desc">%s</div>
    """ % (title)

def newParagraph(newLine):
    return """
    %s
    """ % (newLine)

def Create_Email(content, whoToEmail,whoToCC, subjectLine, attachmentName, newEmailName,nameToggle):
    msg            = MIMEMultipart()
    msg['Subject'] = str(subjectLine)
    msg['To']      = str(whoToEmail)
    if whoToCC != None:
        msg['Cc'] = str(whoToCC)
    # msg['Bcc'] = ""
    msg.add_header('X-Unsent', '1')
    htmlfront = """
    <html>
    <head>
    <style>
    </style>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.0.1/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-+0n0xVW2eSR5OomGNYDnhzAbDsOXxcvSN1TPprVMTNDbiYZCxYbOOl7+AMvyTG2x" crossorigin="anonymous">
    </head>
    <body>
    """
    htmlcontent = content;
    htmlback =  """</body>
    </html>"""
    html=htmlfront+htmlcontent;
    part = MIMEText(html, 'html')
    msg.attach(part)
    outfile_name = desktopPath+'/deliverableEmail.eml'

    if attachmentName!="" and toggleSubmission==True:
        # content+= "<br><br><p>Submission document attached.</p>"

        #this bit of code will add most recent file in downloads to email
        filename = str(attachmentName)
        filename_stripped=filename.replace(filePath,"")
        if nameToggle==True:
            filename_new=filename_stripped
            # filename_new=filename_new.replace(" ", "")
            # file.write("\n name:"+filename_new)
        else:
            filename_new=newEmailName

        filename_to_write=filename_new.replace(" ", "")




        # We assume that the file is in the directory where you run your Python script from
        with open(filePath+filename_new, "rb") as attachment:
            # The content type "application/octet-stream" means that a MIME attachment is a binary file
            attach = MIMEBase("application", "octet-stream")
            attach.set_payload(attachment.read())

        # Encode to base64
        encoders.encode_base64(attach)

        # Add header
        attach.add_header(
            "Content-Disposition",
            f"attachment; filename= {filename_to_write}",
        )

        # Add attachment to your message and convert it to string

        msg.attach(attach)



    with open(outfile_name, 'w+') as outfile:
        gen = generator.Generator(outfile)
        gen.flatten(msg)

    return "complete"

def runScript(projectAttributes,attributesToReceive):
    attributeLength=0
    renameEmail=""
    useOriginalName=True

    useAttributeName=False
    codeInFileName=False
    projectAttributes = ["project.Project.name","code","sg_status_list","project.Project.tags"]
   # projectAttributes+=["project.Project.sg_mainrecipient", "project.Project.sg_ccrecipient", "project.Project.sg_greetingline","project.Project.sg_bodytext","project.Project.sg_listtext","project.Project.sg_closing","project.Project.sg_attachments","project.Project.sg_subjectline"]
    deliverableSpecificAttributes=["sg_ins_deliverable_submission_id"]
    attributesToReceive = "(Shot List:) \t {code}"
    attributesToReceiveRaw=attributesToReceive
    multiLineList=False
    arrayOfListAttributes=[]
    arrayOfFileNameTags=[]

    #sets up communication with Shotgun
    arg = sys.argv[1]
    handler, fullPath = arg.split(":", 1)
    path, fullArgs = fullPath.split("?", 1)
    action = path.strip("/")
    args = fullArgs.split("&")
    get_dict = {}
    for arg in args:
        key, value = map(urllib.parse.unquote, arg.split("=", 1))
        get_dict[key] = value
    file = open(desktopPath+"/log.txt","a+")
    projectAttributes.append(subjectTitleFormatID)
    projectAttributes.append(emailToWhoID)
    projectAttributes.append(ccPeopleID)
    projectAttributes.append(greetingLinePersonID)
    projectAttributes.append(attributesToReceiveID)
    projectAttributes.append(bodyTextID)
    projectAttributes.append(closingTextRawID)
    projectAttributes.append(attachFileToggleID)
    projectAttributes.append(csvHeaderID)
    projectAttributes.append(csvReformatID)
    projectAttributes.append(emailFilenameID)


    # file.write(str(projectAttributes))
    sg=shotgun_api3.Shotgun(
        shotgridURL,
        script_name=scriptName,
        api_key=apiKey,
    )
    post_dict = dict(get_dict)
    entity_ids = post_dict["selected_ids"] or post_dict["ids"]
    entity_ids = [int(id) for id in entity_ids.split(",")]
   # file.write(str(projectAttributes))
  #  projectAttributes.append(subjectTitleFormatID)

    projectInfo = sg.find(
        post_dict["entity_type"],
        [["id", "in", entity_ids]],
        projectAttributes,
    )


    #ATTRIBUTES
    subjectTitleFormat = projectInfo[0][subjectTitleFormatID]
    emailToWho=projectInfo[0][emailToWhoID]
    ccPeople=projectInfo[0][ccPeopleID]
    greetingLinePerson=projectInfo[0][greetingLinePersonID]
    attributesToReceive = projectInfo[0][attributesToReceiveID]
    attributesToReceiveRaw=attributesToReceive
    bodyText = projectInfo[0][bodyTextID]
    closingTextRaw = projectInfo[0][closingTextRawID]
    attachFileToggle=projectInfo[0][attachFileToggleID]
    csvHeaders=projectInfo[0][csvHeaderID]
    csvReformat=projectInfo[0][csvReformatID]
    emailFilename=projectInfo[0][emailFilenameID]


    if "{" in greetingLinePerson and "}" in greetingLinePerson:
        codeExtractArray=greetingLinePerson.split()
        arrayOfAttributeTags=[]
        for i in range(0,len(codeExtractArray)):
            if "{" in codeExtractArray[i] and "}" in codeExtractArray[i]:
                newTag=codeExtractArray[i][codeExtractArray[i].find('{')+1:codeExtractArray[i].find('}')]
                arrayOfAttributeTags.append(newTag)

        greetingLinePersonTag = greetingLinePerson[greetingLinePerson.find('{')+1:greetingLinePerson.find('}')]
        useAttributeName=True
    if emailFilename != None:

        if "{" in emailFilename and "}" in emailFilename:
            codeExtractFilename=emailFilename.split()
            arrayOfFileNameTags=[]
            for i in range(0,len(codeExtractFilename)):
                if "{" in codeExtractFilename[i] and "}" in codeExtractFilename[i]:
                    newFilenameTag=codeExtractFilename[i][codeExtractFilename[i].find('{')+1:codeExtractFilename[i].find('}')]
                    arrayOfFileNameTags.append(newFilenameTag)

            greetingLinePersonTag = greetingLinePerson[greetingLinePerson.find('{')+1:greetingLinePerson.find('}')]
            codeInFileName=True

    if "{" in attributesToReceive and "}" in attributesToReceive:
        codeExtractArrayList=attributesToReceive.split()
        for i in range(0,len(codeExtractArrayList)):
            if "{" in codeExtractArrayList[i] and "}" in codeExtractArrayList[i]:
                newCodeList=codeExtractArrayList[i][codeExtractArrayList[i].find('{')+1:codeExtractArrayList[i].find('}')]
                arrayOfListAttributes.append(newCodeList)
    arrayOfHeaders=[]
    if "(" in attributesToReceive and ")" in attributesToReceive:
        headerText=attributesToReceive.split()
        for i in range(0,len(headerText)):
            if "(" in headerText[i] and ")" in headerText[i]:
                headerList=headerText[i][headerText[i].find('(')+1:headerText[i].find(')')]
                arrayOfHeaders.append(headerList)


    paragraphSplit=[]
    lineSplit=[]
    if "\\n" in attributesToReceiveRaw or "\\t" in attributesToReceiveRaw:
        paragraphSplit=attributesToReceiveRaw.split("\\t")
        for i in range(0,len(paragraphSplit)):
            paragraphSplit[i]=paragraphSplit[i].split("\\n")
            multiLineList=True

            lineSplit=paragraphSplit
    else:
        paragraphSplit=["1"]

    attributesToReceive = arrayOfListAttributes

    bodyContainsAttributes=False
    multiLineBody=False
    arrayOfBodyAttributeTags=[]

    if "{" in bodyText and "}" in bodyText:
        codeExtractArrayBody=bodyText.split()
        for i in range(0,len(codeExtractArrayBody)):
            if "{" in codeExtractArrayBody[i] and "}" in codeExtractArrayBody[i]:
                newTagBody=codeExtractArrayBody[i][codeExtractArrayBody[i].find('{')+1:codeExtractArrayBody[i].find('}')]
                arrayOfBodyAttributeTags.append(newTagBody)

    if "\\n" in bodyText:
        bodyText=bodyText.split("\\n")
        multiLineBody=True

    subjectContainsAttributes=False
    if "{" in subjectTitleFormat and "}" in subjectTitleFormat:
        subjectExtractArray=subjectTitleFormat.split()
        arrayOfSubjectAttributeTags=[]
        for i in range(0,len(subjectExtractArray)):
            if "{" in subjectExtractArray[i] and "}" in subjectExtractArray[i]:
                newSubjectTag=subjectExtractArray[i][subjectExtractArray[i].find('{')+1:subjectExtractArray[i].find('}')]
                arrayOfSubjectAttributeTags.append(newSubjectTag)

        subjectContainsAttributes=True


    multiLineClosing=False
    arrayOfClosingAttributeTags=[]


    if "{" in closingTextRaw and "}" in closingTextRaw:
        codeExtractClosing=closingTextRaw.split()
        for i in range(0,len(codeExtractClosing)):
            if "{" in codeExtractClosing[i] and "}" in codeExtractClosing[i]:
                closingTags=codeExtractClosing[i][codeExtractClosing[i].find('{')+1:codeExtractClosing[i].find('}')]
                arrayOfClosingAttributeTags.append(closingTags)

    if "\\n" in closingTextRaw:
        closingText=closingTextRaw.split("\\n")
        multiLineClosing=True
    formatType = [ "oneLineList"]

    if not os.path.exists(desktopPath+'/submissionDocs'):
        os.makedirs(desktopPath+'/submissionDocs')



    checkFolderEmpty = os.listdir(filePath)
    attributesToReceive+=arrayOfAttributeTags+arrayOfFileNameTags+arrayOfBodyAttributeTags+attributesToReceive+arrayOfSubjectAttributeTags+arrayOfClosingAttributeTags+arrayOfFileNameTags

    # finds selected thingies
    entities = sg.find(
        post_dict["entity_type"],
        [["id", "in", entity_ids]],
        attributesToReceive,
    )
    deliverableSpecific=False
    # projectAttributes=projectAttributes+arrayOfAttributeTags+arrayOfBodyAttributeTags+attributesToReceive+arrayOfSubjectAttributeTags+arrayOfClosingAttributeTags+arrayOfFileNameTags
    # file.write(str(projectAttributes))
    # file.write(str(projectInfo[0][""]))
    if str(entities[0]["type"]) == "CustomEntity26":
        projectAttributes+=deliverableSpecificAttributes
        deliverableSpecific=True

    #AttributeRewrite
    subjectTitleFormat = str(projectInfo[0][subjectTitleFormatID])


# this bit of code uses the tags field to dictate what to put in the deliverable
    tagList=str(projectInfo[0]['project.Project.tags'])
    tagListArray=projectInfo[0]['project.Project.tags']
    tagNamesList=[]
    print(tagListArray)
    listOfTags =[]
    for x in range (0,len(tagListArray)):
        taglistAttributes = []
# https://stackoverflow.com/questions/1679384/converting-dictionary-to-list
        for key, value in tagListArray[x].items():
            temp = [key,value]
            taglistAttributes.append(temp)
        # print(tagListArray)
        listOfTags.append(taglistAttributes)
        # print("index: "+str(x))
        tagNamesList.append(listOfTags[x][1][1])

    html = _html_header()
    if useAttributeName==True:
        greetingMessage=greetingLinePerson.split()
        fullGreetingLine=""
        tagIndex=0
        for i in range(0,len(greetingMessage)):

            if "{" in greetingMessage[i]:
                greetingMessage[i]=entities[0][arrayOfAttributeTags[tagIndex]]
                tagIndex+=1
            fullGreetingLine+=str(greetingMessage[i])
            if i < len(greetingMessage)-1:
                fullGreetingLine+=" "
        html += greetingLine(str(fullGreetingLine.replace("_"," ")))
        # html+=str(desktopPath)

    else:
        html+=greetingLine(greetingLinePerson)

    if codeInFileName==True:
        filenameSplit=emailFilename.split()
        fullNewFileName=""
        fileNameIndex=0
        for i in range(0,len(filenameSplit)):

            if "{" in filenameSplit[i]:
                filenameSplit[i]=entities[0][arrayOfFileNameTags[fileNameIndex]]
                fileNameIndex+=1
            fullNewFileName+=str(filenameSplit[i])
            if i < len(filenameSplit)-1:
                filenameSplit+=" "
        emailFilename=fullNewFileName
        # html += greetingLine(str(fullGreetingLine))
        # html+=str(desktopPath)

    else:
        emailFilename=emailFilename






    if multiLineBody==True:
        attributeIndex=0

        for i in range(0,len(bodyText)):
            splitToWords=bodyText[i].split()
            bodyLine=""
            for i in range(0,len(splitToWords)):
                if "{" in splitToWords[i]:
                    splitToWords[i]=entities[0][arrayOfBodyAttributeTags[attributeIndex]]
                    attributeIndex+=1
                bodyLine+=str(splitToWords[i])
                if i<len(splitToWords)-1:
                    bodyLine+=" "
            if i > 0:
                html+="<br>"
            html+=newParagraph(bodyLine)

    else:
        attributeIndex=0
        splitToWords=bodyText.split()
        bodyLine=""
        for i in range(0,len(splitToWords)):
            if "{" in splitToWords[i]:
                splitToWords[i]=entities[0][arrayOfBodyAttributeTags[attributeIndex]]
                attributeIndex+=1
            bodyLine+=str(splitToWords[i])
            if i<len(splitToWords)-1:
                bodyLine+=" "
        html+=newParagraph(bodyLine)

    valueCounter=0
    runGrouped=False
    for x in range(0,len(entities)):
        valueCounter=0
        values=dict(entities[x])
        for i,j in values.items():
            if i!='type' and i!='id':
                valueCounter+=1
                attributeNumber=len(values)-2
                if attributeNumber==1:
                    runGrouped=True



    if "oneLineList" in formatType:
        # attributeLength
        html+="<p></p>"
        for paragraphs in range(0,len(paragraphSplit)):
            # html+="{"+str(attributeLength)+"}"
            multiLineEntry=False

            if paragraphs>0:
                html+="<br>"

            listTitle=""

            # html+=str(paragraphAttributeArray)+"<br> <br>"
            # html+=str(paragraphAttributeArrayFull)+"<br> <br>"

            # html+=str(paragraphSplit)
            # html+="|"+str(paragraphAttributeArrayFull)+"|"

            if len(paragraphAttributeArrayFull)>0:
                startingValue=len(paragraphAttributeArrayFull)
            else:
                startingValue=0

            #need to find way to get len of this

            for j in range(0,len(paragraphSplit[paragraphs])):
                if "(" in paragraphSplit[paragraphs][j]:
                    listTitle=paragraphSplit[paragraphs][j]
                    listTitle=listTitle.replace("(","")
                    listTitle=listTitle.replace(")","")


                    paragraphSplit[paragraphs].pop(j)


            paragraphAttributes=[]
            additionalAttributes=[]

            # html+=str(paragraphAttributes)+"<br> <br>"

            if (len(paragraphSplit[paragraphs]) > 1):
                multiLineEntry=True
            # html+="{"+str(attributeLength)+"}"
            # attributeLenth+=1
            # html+="start"+str(multiLineEntry)+"<br>"
            html+=listTitle


            for x in range(0,len(entities)):

                incrementListIndex = False

                if multiLineList==True:

                    isImage=False
                    additionalAttributeToggle=False
                    listLineIndex=startingValue+attributeLength

                    # html+=str(listLineIndex)+"|"
                    for i in range (0,len(paragraphSplit[paragraphs])):
                        # if (len(paragraphSplit[paragraphs]) > 1):
                        #     multiLineEntry=True
                        # html+="start"+str(multiLineEntry)
                        lineSplitToWords=paragraphSplit[paragraphs][i].split()

                        fullLine =""

                        #important
                        # html+=str(listLineIndex)
                        # listLineIndex=startingValue
                        # html+=str(listLineIndex)+"|"

                        paragraphAttributes=[]
                        # additionalAttributes=[]
                        # html+="|"+str(i)+"|"
                        if (i > 0):
                            additionalAttributeToggle=True
                        # html+="add?"+str(additionalAttributeToggle)

                        for i in range(0,len(lineSplitToWords)):
                            # html+="|"+str(i)+"|"

                            if "{" in lineSplitToWords[i]:
                                lineSplitToWords[i]=entities[x][arrayOfListAttributes[listLineIndex]]
                                paragraphAttributes.append(arrayOfListAttributes[listLineIndex])
                                incrementListIndex=True

                                if additionalAttributeToggle==True:
                                    # html+="appendAttribute"
                                    if ((arrayOfListAttributes[listLineIndex] in additionalAttributes) == False):
                                        additionalAttributes.append(arrayOfListAttributes[listLineIndex])


                                if arrayOfListAttributes[listLineIndex] == "image":
                                    isImage=True
                                if listLineIndex < len(arrayOfListAttributes)-1:
                                    listLineIndex+=1

                            if isImage==True:
                                fullLine+=_image_html(lineSplitToWords[i],lineSplitToWords[i])
                                isImage=False
                            else:
                                # html+="|"+str("hello")+"|"
                                # fullLine+=str("|"+str(listLineIndex)+"|")

                                fullLine+=str(lineSplitToWords[i])

                                # html+="|"+str(listLineIndex)+"|"
                            if i < len(fullLine)-1:
                                fullLine+=" "
                        # html+="|"+str(paragraphAttributeArrayFull)+"|"
                        # if (i > 0):
                        #     html+="increment"
                        # html+="<br>"
                        if "{" in fullLine:
                            breakdown=fullLine.split("{")
                            for n in range (0,len(breakdown)):
                                # html+=str(breakdown[n])
                                if "}" in breakdown[n]:
                                    breakdown[n]="{"+breakdown[n]
                                    Dict = eval(breakdown[n])
                                    html+=str((Dict.get("name")))

                                else:
                                    html+=str(breakdown[n])


                            # Dict = eval(fullLine)
                            # html+=str((Dict.get("name")))

                        else:
                            html+=(fullLine)

                        html+="<br>"

                    # html+="reset"+str(additionalAttributes)

                        # html+=str(paragraphAttributes)
                    # if (len(paragraphSplit[paragraphs]) > 1):
                    #     multiLineEntry=True
                    # html+="<br>"+str(multiLineEntry)
                else:

                    isImage=False
                    listLine=attributesToReceiveRaw.split()
                    # fullSubjectContent=""
                    fullLine=""
                    listLineIndex=0
                    for i in range(0,len(listLine)):
                        if "{" in listLine[i]:
                            listLine[i]=entities[x][arrayOfListAttributes[listLineIndex]]

                            if arrayOfListAttributes[listLineIndex] == "image":
                                isImage=True

                            listLineIndex+=1
                        # fullLine+=str(i)
                        # fullLine+=str(i)
                        if isImage==True:
                            fullLine+=_image_html(listLine[i],listLine[i])
                        else:
                            fullLine+=str(listLine[i])

                        if i < len(fullLine)-1:
                            fullLine+=" "

                    html+= _shot_name_html(str(fullLine))
            # html+="kachow"
            # html+="kachow"
            # html+=str(listLineIndex)
            # html+="end"+"<br>"
            # html+=str(paragraphAttributes)
            if len(paragraphAttributes) > 0:
                paragraphAttributeArray.append(paragraphAttributes)
                for attributes in range(0,len(paragraphAttributes)):
                    paragraphAttributeArrayFull.append(paragraphAttributes[attributes])
            # attribute=paragraphAttributeArray[0]
            # html+=str(attribute)
            # html+=str(paragraphAttributeArray)
            # html+=str([paragraphAttributeArray])

            if incrementListIndex == True:
                if listLineIndex < len(arrayOfListAttributes)-1:
                    listLineIndex+=1
                    # html+="boom"
            # html+=str(len(additionalAttributes))
                # if listLineIndex < len(arrayOfListAttributes)-1:
                #     listLineIndex+=1
                #     html+="boom"
                # html+="<br>"
                # html+="<br>"
                # html+=str(len(additionalAttributes))+"|"+str(startingValue)
                attributeLength=len(additionalAttributes)
                if attributeLength > 0:
                    attributeLength+=1
        # startingValue+=len(additionalAttributes)

            # html+="|"+str(paragraphAttributeArrayFull)+"|"

        #     html+="<br>"
        # html+="<br>"
            # html+="<br>"+str(multiLineEntry)
        valueCounter=0
        # html+="ouptput"+str(paragraphAttributeArray[0])


    # if runGrouped==False and "groupedList" in formatType:
    #     # html+="<p>Shot List By Item:</p>"
    #     for x in range(0,len(entities)):
    #         valueCounter=0
    #         values=dict(entities[x])
    #         for i,j in values.items():
    #             if i!='type' and i!='id':
    #                 html+=(_shot_name_html(str(j)))
    #                 valueCounter+=1
    #
    #         # html+=(_shot_name_html(str(valueCounter)))
    #         if valueCounter>=2:
    #             html+=("<br>")

    res = defaultdict(list)
    {res[key].append(sub[key]) for sub in entities for key in sub}
    # html+="<p>Shot List:</p>"


    # for x in range(0,len(entities)):
    # html+=_shot_name_html(str(dict(res)))

    groupedValues=dict(res)

    if  "separateLists" in formatType:
        # html+="<br>"
        # html+="Shot List by Type:"
        for i,j in groupedValues.items():
            if i!='type' and i!='id':



                if i=="sg_ins_deliverable_long_name":
                    i="Shot List"
                else:
                    i=str(i.replace("sg_ins_",""))
                    i=str(i.replace("sg_",""))
                    i=str(i.replace("_"," "))
                    i=string.capwords(i)




                html+=("<b>"+str(i)+":"+"</b>")
                html+="<br>"
                for x in range(0,len(j)):
                    if i == "Image":
                        html+=_image_html(j[x],j[x])
                    else:
                        html+=_shot_name_html(j[x])
            html+="<br>"

                # html+=(_shot_name_html(str(j)))
            # html+=("<br>")
        # values=dict(entities[x])
        # for i,j in values.items():
        #     if i!='type' and i!='id':
        #         html+=(_shot_name_html(str(j)))
        # html+=("<br>")


    # Generates to: field in email from Project attribute
    # emailToWho = ""
    # emailArray = str(((entities[0]['project.Project.sg_mainemailrecipient']))).split(",")
    # for x in range(0,len(emailArray)):
    #     emailToWho+=str(emailArray[x]);
    #     if (x<len(emailArray)-1):
    #         emailToWho+=", "
    if attachFileToggle==True:
        #resumehere
        html+="<br>Submission document attached.<br>"

    closingTextFinal=""
    # html+="<br>"
    if multiLineClosing==True:
        closingAttributeIndex=0

        for i in range(0,len(closingText)):
            closingTextTemp=closingText[i].split()
            closingTextFinal=""
            for i in range(0,len(closingTextTemp)):
                if "{" in closingTextTemp[i]:
                    closingTextTemp[i]=entities[0][arrayOfClosingAttributeTags[closingAttributeIndex]]
                    closingAttributeIndex+=1
                closingTextFinal+=str(closingTextTemp[i])
                if i<len(closingTextTemp)-1:
                    closingTextFinal+=" "
            html+="<br>"
            html+=(closingTextFinal)
        if attachFileToggle == True:
            html+="<br>"
        # html+="<br>"
    else:
        closingAttributeIndex=0
        closingTextTemp=closingTextRaw.split()
        bodyLine=""
        for i in range(0,len(closingTextTemp)):
            if "{" in closingTextTemp[i]:
                closingTextTemp[i]=entities[0][arrayOfClosingAttributeTags[closingAttributeIndex]]
                closingAttributeIndex+=1
            closingTextFinal+=str(closingTextTemp[i])
            if i<len(closingTextTemp)-1:
                closingTextFinal+=" "
        html+="<br>"
        html+=closingTextFinal
        if attachFileToggle == True:
            html+="<br>"
        # html+=str(os.getcwd())
    # contentToWrite = html + str(closingTextFinal)+ "<br><p>Thank you,</p><p>INS</p>" +"</body></html>"

    contentToWrite = html +"</body></html>"

    #delete this line eventually
    # emailToWho="ryryk1@gmail.com"
    # ccPeople="kawahara@gmail.com"


    # email subject part
    fullSubjectContent=""
    if subjectContainsAttributes==True:
        subjectContent=subjectTitleFormat.split()
        # fullSubjectContent=""
        subjectTagIndex=0
        for i in range(0,len(subjectContent)):

            if "{" in subjectContent[i]:
                subjectContent[i]=entities[0][arrayOfSubjectAttributeTags[subjectTagIndex]]
                subjectTagIndex+=1
            fullSubjectContent+=str(subjectContent[i])
            if i < len(subjectContent)-1:
                fullSubjectContent+=" "
        emailSubject = str(fullSubjectContent)

        # html += greetingLine(str(arrayOfBodyAttributeTags))
    else:
        fullSubjectContent=subjectTitleFormat
        emailSubject=(fullSubjectContent)


    # submission_ID=projectInfo[0]["sg_ins_deliverable_submission_id"]
    # emailSubject=submission_ID


    # latest_file=""
    if len(checkFolderEmpty)==0 or attachFileToggle==False:
        latest_file=""

    else:
    # https://stackoverflow.com/questions/39327032/how-to-get-the-latest-file-in-a-folder
        list_of_files = glob.glob(filePath+'*') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getctime)

        useOriginalName=True
        # file.write("\n"+latest_file.replace(filePath,""))
        #rename here
        # renameEmail=emailFilename+str(split_tup[1])

        #manipulates CSV
        split_tup = os.path.splitext(latest_file)
        split_tup=list(split_tup)

        if emailFilename==None:
            useOriginalName=True

        else:
            renameEmail=emailFilename+split_tup[1]
            useOriginalName=False


        if (split_tup[1]==".csv") and (csvReformat==True):
            data = pd.read_csv(latest_file)


            if csvHeaders != None:
                csvHeaders=csvHeaders.split(",")
                # data.pop(data.columns.values[-1])

                    # raise Exception("Headers field has more attributes than CSV.")
                for item in range(0,len(csvHeaders)):
                    csvHeaders[item]=csvHeaders[item].strip()

            if 'Id'in data.columns:
                data.pop('Id')
            if 'Status'in data.columns:
                data.pop('Status')
            if 'Project'in data.columns:
                data.pop('Project')
            if "Unnamed: 16" in data.columns:
                data.pop("Unnamed: 16")
            if 'Date Updated' in data.columns:
                date_split = data['Date Updated'].str.split()
                data['Date Updated']=(date_split.str[0:1] + date_split.str[2:]).str.join(' ')
            if csvHeaders != None:
                data.columns=[csvHeaders]
            # data['Date Updated']=(date_split.str[0:1] + date_split.str[2:]).str.join(' ')
            data.to_csv(latest_file, encoding='utf-8',index=False)
        if useOriginalName==False:
            os.rename(latest_file, filePath+renameEmail)
        else:
            renameEmail=latest_file
            file.write(str(latest_file))



    #this line creates the email
    Create_Email(contentToWrite,emailToWho,ccPeople,emailSubject,latest_file,renameEmail,useOriginalName)

    #logs the successful implementation in the log file
    now = datetime.now()
    todayTime = now.strftime("%m/%d/%Y %H:%M:%S")
    # if deliverableSpecific==True:
    #     submission_ID=projectInfo[0]["sg_ins_deliverable_submission_id"]
    #
    #     completeMessage="\n"+projectInfo[0]["project.Project.name"]+" | "+str(submission_ID)+" "+"successfully completed on: "+str(todayTime)
    # else:
    completeMessage="\n"+projectInfo[0]["project.Project.name"]+" | "+"successfully completed on: "+str(todayTime)
    file.write(completeMessage)

    # file.write(str(entities[0]["type"]))


    file.close()
    emailDraft=desktopPath+'/deliverableEmail.eml'
    #opens up the email draft
    subprocess.call(["open", emailDraft])
    #triggers a system notification
    body=completeMessage
    os.system("osascript -e 'display notification \""+body+"\" with title \"Email Generator Script\"'")
    # os.system("osascript -e 'Tell application \"System Events\" to display dialog \""+body+"\"'")


#sets up error notifier
logging.basicConfig(
    filename=desktopPath+"/log.txt",
    filemode='a+'
)

try:
    runScript(projectAttributes,attributesToReceive)
#notifies user of error via pop-up box and then opens log
except Exception as Argument:
    errormessage="Error occurred. Please check log."
    logging.error("Exception occurred", exc_info=True)
    os.system("osascript -e 'Tell application \"System Events\" to display dialog \""+errormessage+"\"'")
    subprocess.call(["open", "-R", desktopPath+"/log.txt"])



# https://realpython.com/python-logging/

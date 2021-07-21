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

#this path is where the .eml template and log will be sent
desktopPath = os.path.dirname(__file__)
downloadsPath= "~/Downloads/"
filePath = desktopPath + "/submissionDocs/"
projectAttributes = []
attributesToReceive = []
toggleSubmission=True
incrementListIndex=False
arrayOfLineswithAttributes = []
paragraphAttributes=[]
paragraphAttributeArray=[]
startingValue=0




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
<body>
"""

def greetingLine(projectName):
    return """
    <p> %s </p>
    """ % (projectName)

def _image_html(link, image):
    return """
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
    <p>%s</p>
    """ % (newLine)

def Create_Email(content, whoToEmail,whoToCC, subjectLine, attachmentName):
    msg            = MIMEMultipart()
    msg['Subject'] = str(subjectLine)
    msg['To']      = str(whoToEmail)
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
    outfile_name = desktopPath+'deliverableEmail.eml'

    if attachmentName!="" and toggleSubmission==True:
        # content+= "<br><br><p>Submission document attached.</p>"

        #this bit of code will add most recent file in downloads to email
        filename = str(attachmentName)
        filename_stripped=filename.replace(filePath,"")

        # We assume that the file is in the directory where you run your Python script from
        with open(filename, "rb") as attachment:
            # The content type "application/octet-stream" means that a MIME attachment is a binary file
            attach = MIMEBase("application", "octet-stream")
            attach.set_payload(attachment.read())

        # Encode to base64
        encoders.encode_base64(attach)

        # Add header
        attach.add_header(
            "Content-Disposition",
            f"attachment; filename= {filename_stripped}",
        )

        # Add attachment to your message and convert it to string

        msg.attach(attach)



    with open(outfile_name, 'w+') as outfile:
        gen = generator.Generator(outfile)
        gen.flatten(msg)

    return "complete"

def runScript(projectAttributes,attributesToReceive):


    useAttributeName=False
    projectAttributes = ["project.Project.name","code","sg_status_list","project.Project.tags"]
    projectAttributes+=["project.Project.sg_mainrecipient", "project.Project.sg_ccrecipient", "project.Project.sg_greetingline","project.Project.sg_bodytext","project.Project.sg_listtext","project.Project.sg_closing","project.Project.sg_attachments","project.Project.sg_subjectline"]
    deliverableSpecificAttributes=["sg_ins_deliverable_submission_id"]
    attributesToReceive = "(Shot List:) \t {code}"
    attributesToReceiveRaw=attributesToReceive
    multiLineList=False
    arrayOfListAttributes=[]

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
    file = open(desktopPath+"log.txt","a+")
    sg=shotgun_api3.Shotgun(
        "https://ryantoolstesting.shotgrid.autodesk.com/",
        script_name="EmailGenerator",
        api_key="wiarjthh7%hhwfajlybsjrDxb",
    )
    post_dict = dict(get_dict)
    entity_ids = post_dict["selected_ids"] or post_dict["ids"]
    entity_ids = [int(id) for id in entity_ids.split(",")]

    projectInfo = sg.find(
        post_dict["entity_type"],
        [["id", "in", entity_ids]],
        projectAttributes,
    )


    #ATTRIBUTES
    subjectTitleFormat = projectInfo[0]["project.Project.sg_subjectline"]
    emailToWho=projectInfo[0]["project.Project.sg_mainrecipient"]
    ccPeople=projectInfo[0]["project.Project.sg_ccrecipient"]
    greetingLinePerson=projectInfo[0]["project.Project.sg_greetingline"]
    attributesToReceive = projectInfo[0]["project.Project.sg_listtext"]
    attributesToReceiveRaw=attributesToReceive
    bodyText = projectInfo[0]["project.Project.sg_bodytext"]
    closingTextRaw = projectInfo[0]["project.Project.sg_closing"]
    attachFileToggle=projectInfo[0]["project.Project.sg_attachments"]

    if "{" in greetingLinePerson and "}" in greetingLinePerson:
        codeExtractArray=greetingLinePerson.split()
        arrayOfAttributeTags=[]
        for i in range(0,len(codeExtractArray)):
            if "{" in codeExtractArray[i] and "}" in codeExtractArray[i]:
                newTag=codeExtractArray[i][codeExtractArray[i].find('{')+1:codeExtractArray[i].find('}')]
                arrayOfAttributeTags.append(newTag)

        greetingLinePersonTag = greetingLinePerson[greetingLinePerson.find('{')+1:greetingLinePerson.find('}')]
        useAttributeName=True

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

    # finds selected thingies
    entities = sg.find(
        post_dict["entity_type"],
        [["id", "in", entity_ids]],
        attributesToReceive,
    )
    deliverableSpecific=False
    projectAttributes=projectAttributes+arrayOfAttributeTags+arrayOfBodyAttributeTags+attributesToReceive+arrayOfSubjectAttributeTags+arrayOfClosingAttributeTags
    if str(entities[0]["type"]) == "CustomEntity26":
        projectAttributes+=deliverableSpecificAttributes
        deliverableSpecific=True

    #AttributeRewrite
    subjectTitleFormat = str(projectInfo[0]["project.Project.sg_subjectline"])


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
                greetingMessage[i]=projectInfo[0][arrayOfAttributeTags[tagIndex]]
                tagIndex+=1
            fullGreetingLine+=str(greetingMessage[i])
            if i < len(greetingMessage)-1:
                fullGreetingLine+=" "
        html += greetingLine(str(fullGreetingLine))

    else:
        html+=greetingLine(greetingLinePerson)

    if multiLineBody==True:
        attributeIndex=0

        for i in range(0,len(bodyText)):
            splitToWords=bodyText[i].split()
            bodyLine=""
            for i in range(0,len(splitToWords)):
                if "{" in splitToWords[i]:
                    splitToWords[i]=projectInfo[0][arrayOfBodyAttributeTags[attributeIndex]]
                    attributeIndex+=1
                bodyLine+=str(splitToWords[i])
                if i<len(splitToWords)-1:
                    bodyLine+=" "
            html+=newParagraph(bodyLine)

    else:
        attributeIndex=0
        splitToWords=bodyText.split()
        bodyLine=""
        for i in range(0,len(splitToWords)):
            if "{" in splitToWords[i]:
                splitToWords[i]=projectInfo[0][arrayOfBodyAttributeTags[attributeIndex]]
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

        for paragraphs in range(0,len(paragraphSplit)):

            if paragraphs>0:
                html+="<br>"

            listTitle=""

            if len(paragraphAttributeArray)>0:
                startingValue=len(paragraphAttributeArray[0])
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

            html+=listTitle


            for x in range(0,len(entities)):
                incrementListIndex = False

                if multiLineList==True:

                    isImage=False

                    for i in range (0,len(paragraphSplit[paragraphs])):

                        lineSplitToWords=paragraphSplit[paragraphs][i].split()

                        fullLine =""

                        #important
                        listLineIndex=startingValue
                        paragraphAttributes=[]
                        for i in range(0,len(lineSplitToWords)):
                            if "{" in lineSplitToWords[i]:
                                lineSplitToWords[i]=entities[x][arrayOfListAttributes[listLineIndex]]
                                paragraphAttributes.append(arrayOfListAttributes[listLineIndex])
                                incrementListIndex=True


                                if arrayOfListAttributes[listLineIndex] == "image":
                                    isImage=True
                                if listLineIndex < len(arrayOfListAttributes)-1:
                                    listLineIndex+=1

                            if isImage==True:
                                fullLine+=_image_html(lineSplitToWords[i],lineSplitToWords[i])
                            else:
                                fullLine+=str(lineSplitToWords[i])
                            if i < len(fullLine)-1:
                                fullLine+=" "
                        html+=(fullLine+"<br>")

                else:
                    file.write("hello")

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
          
            # html+=str(listLineIndex)
            # html+="end"+"<br>"
            # html+=str(paragraphAttributes)
            if len(paragraphAttributes) > 0:
                paragraphAttributeArray.append(paragraphAttributes)
            # attribute=paragraphAttributeArray[0]
            # html+=str(attribute)
            # html+=str(paragraphAttributeArray)
            # html+=str([paragraphAttributeArray])

            if incrementListIndex == True:
                if listLineIndex < len(arrayOfListAttributes)-1:
                    listLineIndex+=1
                  

                # if listLineIndex < len(arrayOfListAttributes)-1:
                #     listLineIndex+=1
                
                # html+="<br>"


        #     html+="<br>"
        # html+="<br>"
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
                    closingTextTemp[i]=projectInfo[0][arrayOfClosingAttributeTags[closingAttributeIndex]]
                    closingAttributeIndex+=1
                closingTextFinal+=str(closingTextTemp[i])
                if i<len(closingTextTemp)-1:
                    closingTextFinal+=" "
            html+="<br>"
            html+=(closingTextFinal)
    else:
        closingAttributeIndex=0
        closingTextTemp=closingTextRaw.split()
        bodyLine=""
        for i in range(0,len(closingTextTemp)):
            if "{" in closingTextTemp[i]:
                closingTextTemp[i]=projectInfo[0][arrayOfClosingAttributeTags[closingAttributeIndex]]
                closingAttributeIndex+=1
            closingTextFinal+=str(closingTextTemp[i])
            if i<len(closingTextTemp)-1:
                closingTextFinal+=" "
        html+="<br>"
        html+=closingTextFinal
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
                subjectContent[i]=projectInfo[0][arrayOfSubjectAttributeTags[subjectTagIndex]]
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
        file.write("\n"+latest_file.replace(filePath,""))

    #this line creates the email
    Create_Email(contentToWrite,emailToWho,ccPeople,emailSubject,latest_file)

    #logs the successful implementation in the log file
    now = datetime.now()
    todayTime = now.strftime("%m/%d/%Y %H:%M:%S")
    if deliverableSpecific==True:
        submission_ID=projectInfo[0]["sg_ins_deliverable_submission_id"]

        completeMessage="\n"+projectInfo[0]["project.Project.name"]+" | "+str(submission_ID)+" "+"successfully completed on: "+str(todayTime)
    else:
        completeMessage="\n"+projectInfo[0]["project.Project.name"]+" | "+"successfully completed on: "+str(todayTime)
    file.write(completeMessage)

    # file.write(str(entities[0]["type"]))


    file.close()
    emailDraft=desktopPath+'deliverableEmail.eml'
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

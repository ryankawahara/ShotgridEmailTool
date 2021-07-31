# Deliverable Email Generator
by Ryan Kawahara

## Overview
This script gathers data from Shotgrid using the Shotgrid API (https://developer.shotgridsoftware.com/python-api/installation.html) and generates an email (.eml) that can be opened in Apple Mail.

It also generates a log.txt file that can be used for debugging.

Only compatible with MacOS. Needs to have Python3 installed.

## Installation

After you've downloaded this repo, locate the shell file called "deliverableEmailSetup.sh"

Then, open Terminal and execute the following lines. Replace #pathtodownloadedfolder with the path to the downloaded folder.

```
cd #pathtodownloadedfolder
sh deliverableEmailSetup.sh
```
The script will ask to open an app called "emailScript2." Follow the Mac prompts to authenticate.

## Setting Up Shotgrid
Once the setup script is complete, you can set up the infrastructure on the Shotgrid side.
- Create a "Script," and put the name and script key into the emailSetup.py (~/Documents/deliverableEmailTool/emailSetup.py)
- Create an "Action Menu Item." Call it whatever you'd like, and then enter this for the URL: emailScript2://emailScript?runCode
- In the "Projects" view, click Fields > Manage Project Fields > Add a New Field
  - Create the following project fields. Below are what they will be used for, but you can call them whatever you'd like
    - subjectline
    - mainrecipient
    - ccrecipient
    - greetingline
    - bodytext
    - listtext
    - closing
    - attachments (should be a checkbox)
  - With those fields created, you can create a widget in the Project Details (Overview) page like the one shown below.
    
    - <img width="273" alt="Screen Shot 2021-07-27 at 8 08 01 PM" src="https://user-images.githubusercontent.com/64173139/127257621-6cfc5d74-9d3d-42cd-a3bb-0afca83c8efc.png">

  - That input creates the following email as a .eml file, which you can move into your drafts folder
    - <img width="989" alt="Screen Shot 2021-07-27 at 8 10 18 PM" src="https://user-images.githubusercontent.com/64173139/127257696-2dddec7d-75cb-4849-8bc4-60c4c50330e1.png"> 
    - <img width="279" alt="Screen Shot 2021-07-27 at 8 22 20 PM" src="https://user-images.githubusercontent.com/64173139/127258593-487c8eda-d100-41c6-8194-bfe269c03a03.png">

 


## Syntax

- For the main recipient and cc recipient fields, enter emails separated by a comma (without a space).
  - Ex: email1@gmail.com,email2@gmail.com
- For all other fields, you can use {} syntax.
  - For the subject, greeting, body, and closing use the {} syntax with project attributes.
    - Project attributes must be preceded by project.Project.
    - Ex: {project.Project.Name} yields the project name.
  - For everything except the greeting line, you can use \n to go to a new line.
  - For the list section, you can use \t to make new lists.
  - Use {} syntax to get attributes of the selected items. Ex: {code}
  - Putting a word in parentheses makes it a header.
    - Ex: (IDs) \t {id} yields the following 
      - <img width="49" alt="Screen Shot 2021-07-27 at 8 24 20 PM" src="https://user-images.githubusercontent.com/64173139/127258727-59fda68b-94e2-47f6-9426-c3b997f76543.png">




## Things to Avoid

- Do not put a comma directly after the curly brackets, as it will disappear.
  - Put on space in between. Ex: {code} , {id}
- You can't put two \t right after one another. Ex: \t \t 
- Avoid attributes with multiple components like Assets.




###### Created using this tutorial from Shotgrid: https://www.youtube.com/watch?v=u3u8pL9o-as

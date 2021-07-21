# ShotgridEmailTool
Creates email template from selected Shotgrid items. Fully customizable.


This script gathers data from Shotgrid using the Shotgrid API (https://developer.shotgridsoftware.com/python-api/installation.html) and generates an email (.eml) that can be opened in Apple Mail.

It also generates a log.txt file that can be used for debugging.

To run this properly, you will need:

Python 3 (https://www.python.org/downloads/)
Pip (https://pip.pypa.io/en/stable/installing/)
Virtualenv (https://packaging.python.org/guides/installing-using-pip-and-virtual-environments/)
Create a venv called email_virtualenv by executing the following code in Terminal (mac)
'''
cd Documents; python3 -m venv email_virtualenv
'''
Activate the email_virtualenv by executing this line in terminal:

'''
source ~/Documents/email_virtualenv/bin/activate
'''
then execute this line:
'''
pip install git+git://github.com/shotgunsoftware/python-api.git
'''
This will install the Shotgun API

If you get an xcrun error, follow the instructions on this page: https://flaviocopes.com/fix-xcrun-error-invalid-active-developer-path/
To set up, drag the emailScript2 into your Applications folder on Mac. Then, to run it, navigate to Shotgrid and, when items are selected, click "more" and find "emailApp." Click it to run the script.

Created using this tutorial from Shotgrid: https://www.youtube.com/watch?v=u3u8pL9o-as

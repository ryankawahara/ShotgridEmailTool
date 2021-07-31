#!/usr/bin/env bash
# cp -R /Volumes/Chwumbo/engineering/emailScript2.app /Applications/

cd "$(dirname "$0")"
python3 -m pip install --user --upgrade pip

mkdir -p ~/Documents/deliverableEmailTool
mkdir -p ~/Documents/deliverableEmailTool/submissionDocs
install -v -m 755 emailScript.py ~/Documents/deliverableEmailTool
install -v -m 755 emailSetup.py ~/Documents/deliverableEmailTool
cp -r emailScript2.app ~/Documents/deliverableEmailTool

python3 -m pip install --user virtualenv
cd
cd ~/Documents/deliverableEmailTool; python3 -m venv email_virtualenv
source ~/Documents/deliverableEmailTool/email_virtualenv/bin/activate
pip install git+git://github.com/shotgunsoftware/python-api.git
pip install pandas
open -a emailScript2

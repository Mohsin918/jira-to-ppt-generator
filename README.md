# UseCase 1

This script generates a PowerPoint presentation of the PI2023.2/RD5 Executive Summary from JIRA issues.

## Requirements

To run the script, you need the following libraries installed:

- pptx
- jira
- configparser
- itertools
- pyexpat

You can install them using pip:

pip install python-pptx jira configparser itertools pyexpat

# UseCase 2

This script generates a PowerPoint presentation of the Adoption to Retention roadmap from JIRA issues.

## Requirements

To run the script, you need the following libraries installed:

- collections
- pptx
- jira
- configparser
- tqdm

You can install them using pip:

pip install python-pptx jira configparser tqdm

# Configuration

Create a config.ini file in the same folder as your script with the following content:

[jira]
server = your_jira_server_url
username = your_jira_username
password = your_jira_password

Replace your_jira_server_url, your_jira_username, and your_jira_password with your JIRA server URL, username, and password, respectively.
In this case, server = https://jira.tools.sap

To run the script, open a terminal in the folder containing the script and execute the following command:

python UseCase2.py
python UseCase1.py

## To RUN UI

cf login
cf push jira-to-ppt
https://jira-to-ppt.cfapps.eu10.hana.ondemand.com/
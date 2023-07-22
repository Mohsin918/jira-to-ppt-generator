# JIRA-TO-PPT-GENERATOR

JIRA-TO-PPT is a Python script that uses the JIRA API to convert all issues that are in the "In Progress", "Blocked", or "Ready for Review" state to a PowerPoint presentation. The presentation is sorted by quarter, and each issue is represented by a slide that includes the issue's title, description, status, and assignee.

The script is also built on a Flask web app that allows users to specify the JIRA project and the quarter that they want to generate the presentation for. The app then runs the script and generates the presentation, which can be downloaded by the user.

This project is useful for teams that want to track the progress of their JIRA issues in a visual format. The PowerPoint presentation can be used to share the status of the issues with stakeholders or to track the progress of the team over time.

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

###Screenshots

![one](https://github.com/Mohsin918/jira-to-ppt-generator/assets/58115232/d2296d12-ad01-4982-a666-f2f46c4d0332)

![two](https://github.com/Mohsin918/jira-to-ppt-generator/assets/58115232/bfaabbfa-a178-4525-99fc-d1ab29838cae)


![three](https://github.com/Mohsin918/jira-to-ppt-generator/assets/58115232/ecc2cbb9-1905-41cb-890b-6b128100342f)

![four](https://github.com/Mohsin918/jira-to-ppt-generator/assets/58115232/f96bef72-36c7-491a-9c84-a634f30e8bd6)



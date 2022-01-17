<h1>Webex_data_generate</h1>

<h2>How to Install and Run This App</h2>

<b>STEP 1</b>: Make sure Python is installed on your workstation. If you dont have it, you can get it here ---> https://www.python.org/downloads/

<b>STEP 2</b>: Fork this repo. Start by making sure you're logged into GitHub. Then click the Fork button in the upper-right hand corner of this page (https://github.com/vignesh271121/Webex_data_generate) and follow the prompts.

<b>STEP 3</b>: Clone this repo. Click the green code button and copy the URL listed under 'HTTPS'. Now go to you IDE, such as VS Code, PyCharm or Atom, find a place to clone it and type 'git clone' plus the URL you just copied. For example 'git clone https://github.com/vignesh271121/Webex_data_generate.git'

<b>STEP 4</b>: Create a virtual environment. cd into the Webex_data_generate folder. Type 'python -m venv venv' and then 'source venv/bin/activate' for Mac and Linix or 'source venv/scripts/activate' on Windows. You'll know it worked when you see '(venv)' at the beginning of your command prompt.

<b>STEP 5</b>: Install the requirements. Type pip intall -r requirements.txt

<b>STEP 6</b>: Edit the download location. Open the file report_App/report_download.py and on around line 47, replace 'Enter your Path folder' with the folder of your choice for the Excel report to be placed. Make sure that folder is created and present with full path.

<b>STEP 7</b>: Run the app. From the Webex_data_generate folder, run the command 'python manage.py runserver' and use your web browser go to the URL presented in the terminal, such as http://127.0.0.1:8000/. You'll find your authentication token here (https://developer.webex.com/docs/getting-started). Choose a date and room and hit 'Download'. Your results will print.

![image](https://user-images.githubusercontent.com/97229745/148540814-6a4bd522-90b7-4a7f-9baf-58d2fb00bd46.png)


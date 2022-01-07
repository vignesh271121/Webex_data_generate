Before cloning the repo setup the environment using the following command.
#pip install requirements.txt (The file is attached)
Next step is to  clone the repo and open the project folder using any IDE like pycharm or Atom Editor. Then edit the 50th line with the folder name were we want excel report to be saved. (Make sure that folder is created and present with full path – Example  - (Mention the folder path))
 
Save the file and runserver from the location of project folder were manage.py resides.
#python manage.py runserver
Now we can see the below screen
 
Here for Authentication Token navigate to -[https://developer.webex.com/docs/getting-started] and click on copy icon
 
Next paste that token which is valid for 12 hours and download the chat messages with details.
Note – There is a small bug  - If we select Nov Month we are getting only oct month data. Please check


![image](https://user-images.githubusercontent.com/97229745/148540814-6a4bd522-90b7-4a7f-9baf-58d2fb00bd46.png)

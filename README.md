Setup Instructions:
To make this script work, you'll need to download a few packages using Python/Conda first.
Conda allows you to download packages from Python easily. You can download Conda here:
https://www.anaconda.com/distribution/
Scroll to the bottom, make sure that you are on the Windows tab, and choose Download (I used 64-bit, since that is what our Windows computer runs on.)
Run the exe file that is downloaded. When the installer says 'Choose Install Location' at the top, make a note of the Destination Folder Path that it gives you, next to the 'Browse' button. For instance, my path says 'C:\Users\LeeHouse\Anaconda3'. 
Search for 'Anaconda Prompt' in the search bar and then open up Anaconda, which should look like a small black box. next, run: 
pip install pywin32==224 bs4 selenium comtypes==1.1.4 xlwings requests
To install all packages. 
If all of this runs well, this is all you'll have to do to set up the python part of this script.

The next step will be to set up Excel. This script is designed to work for any excel sheets (that are similarily formatted to the one you sent me), but each excel sheet must be saved as an .xlsm (macro-enabled worksheet). 
First, you must enable the developer tab if you haven't already, by going to File->Options->Customize Ribbon->Developer Checkbox. 
Next, we'll need xlwings to communicate between excel and the code. You can download the xlwings add in here:
https://github.com/xlwings/xlwings/releases
Click on the top most xlwings.xlam (at the moment, version 0.17.0) to download it. 
Then, go to the new Developer tab, click add ins, and then navigate to this xlwings.xlam folder that you just downloaded. When you click on this it should add an xlwings tab to your Excel spreadsheet, next to the Developer tab. 
In the xlwings tab, there should be an entry for Interpreter. Here, put in the path that you got from the Anaconda installer, then add \python.exe at the end. For some reason, mine only worked when I put two backlashes between everything, so my path became:
C:\\Users\\LeeHouse\\Anaconda3\\python.exe
This tells xlwings to use the python that lives here, where we have downloaded all of our packages.

Finally, we can set up our Visual Basic code! Go back to the Developer tab and then click 'Visual Basic'. Go to Tools->References and make sure that xlwings is checked and anything that has MISSING in front of it is unchecked. If you can't find xlwings, go to Browse and look for the xlwings.xlam file you downloaded earlier. 
Click OK and then find your project in the left menu. Right click and then Insert->Module. You can name it something like 'scraper' by clicking on the title in the bottom left menu. In the file, paste this code:
Option Explicit
Sub RunScraper()
RunPython ("import scraper; scraper.main()")
End Sub
Finally, you can set up your actual excel sheet. As of right now, the code looks at 5 specific cells to gather information: Four are for the dates that you want it to look for, set at the current correct dates. You can change them as needed, but you must put them in the exact format the website lists them as (i.e. MON ##) and put them in quotes (otherwise excel will automatically format it in its own date format). These dates must be in cells G1, G2, H2, and I2 specifically. 
Finally, the last step to make this work is link to the correct chromedriver. I've included two chromedrivers in the same folder, one for versions 79 and 80. You can check what version you have by going to chrome://downloads. 
The script won't work if the chromedriver won't match the chrome version, so if you have an older version of Chrome you can find one your own version at https://chromedriver.chromium.org/downloads (or I can send one for you!)
Anyway, you'll need to put the path to the correct chromedriver in the box J2, which basically entails finding the right chromedriver file in the windows file manager and then clicking the top bar to copy the path, then adding \chromedriver## at the end. For example, the path I have in my excel sheet is C:\Users\LeeHouse\Desktop\isabel excel stuff\webscraper_project\chromedriver79, but obviously it will be different for different computers.
You can add a button that will trigger the webscraper by going to Developer->Insert->Form Control->Button (first one), clicking where you want this button, and then clicking the RunScraper module. 
When you click this button, the script should start running! A Chrome window will pop up and the software will redirect to the pages needed (you don't need to do anything at this point). 

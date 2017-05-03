# Build Your First Office Integration

## Overview
With Office Add-ins, you can add new functionality to Office and embed your rich and interactive content. In this walk through, you will get hands-on experience developing an Excel add-in with Yo Office. Your final add-in will be able to write data to a workbook and apply basic formatting, as well as bound a chart to that data. If you would instead like to create an add-in using Visual Studio, see Building Your First Office Integration with Visual Studio.

## Tools You’ll Use
Command Prompt
Yo Office
Visual Studio Code

## Part 1: Getting Started

1.	Search for and open the command prompt. 
![]() 

2.	Run the following command to install Yo Office and its dependencies: 
```
npm install -g yo generator-office
```
 
3.	Use the following commands to change the directory to the Documents folder, then create a folder called “myAddin”:
```
# Change directory
cd Documents

# Create new folder
mkdir myAddin
```
 
4.	Change the directory to the myAddin folder. Then, run the Office Yeoman generator to create the project scaffolding, using the following commands: 

 ```
 # Change directory
 cd myAddin
 
 # Run Yo Office
 yo office
 ```
 
  On first launch, you may be prompted to anonymously report usage statistics. This allows the Yeoman team to continue to improve the platform and allow us to deliver the best services. We recommend allowing your data to be collected, but if you don’t feel comfortable sharing your data, feel free to decline. Once past this, Yo Office will launch and prompt you with the following questions: 
  
  |Question|Answer|
  |-|-|
  |Would you like to create a new subfolder for your project?|No <br> (Press Enter to accept the default)|
  |What do you want to name your add-in?|My Office Add-in <br> (Press Enter to accept the default)|
  |Which Office client application would you like to support?|Excel <br> (Press Enter to accept the default)|
  |Would you like to create a new add-in?|Yes, I want a new web app and manifest <br> (Press Enter to accept the default)|
  |Would you like to use TypeScript?|Yes <br> (Press Enter to accept the default)|
  |Choose a framework:|Jquery <br> (Press Enter to accept the default)|
 
  Once you have entered the above information, Yo Office will prompt you to open a resource page for more information and guidance. When developing an Office add-in on your own, the resource page provides a useful guide for the various stages in the Office development process. Since this quickstart will guide you through the process of creating an add-in, feel free to decline opening the page (type “n”, then press Enter). However, if you are curious about the available information, then choose Yes (press Enter to accept the default).
  
  ![Yo Office Screenshot]()

 
  Now, Yo Office will create your templates and install any remaining dependencies necessary for building the rest of your add-in. This may take a couple minutes. 

5.	Host your add-in. Once Yo Office has finished running, you can host your add-in locally, or use any web server or hosting technology – just make sure that the add-in is served using HTTPS, and update the add-in’s source location in the manifest. For this quick start, host the add-in through npm using the following command:
```
npm start
```
  
 Once successfully started, the following page will open in your default browser.
![Screenshot of default template in browser]()

Typically, if this is the first time an Office Add-in is hosted on a machine in this way, the browser will throw an error, and you will need to add the self-signed security certificate that is created as a trusted root certificate or your add-in will not display. This guide bypasses that step.

6.	Load the add-in into Office. The easiest way to do this is by sideloading the add-in in Office Online:
  
    a.	Click here to go to Excel Online and create a blank workbook. You will be required to sign in with your work, school, or Microsoft account.
  
    b.	Go to **Insert > Office Add-ins**
  
    c.	On the My Add-ins tab (or My Organization tab if you're signed in to a work or school account), you will see a link in the upper-right corner of the dialog box to **Upload My Add-in** or **Manage My Add-ins**. Manage My Add-ins will open a menu where you can then choose Upload My Add-in.
  
    d.	In the Upload Add-in dialog, choose **Browse** and select the my-office-add-in-manifest.xml file from the “myAddin” folder in Documents. Then, choose **Upload**. Your add-in will load in Excel Online.

![Screenshot of show taskpane button]()

![Screenshot of default add-in template in Office]()

## Part 2: Customize the Office Ribbon UI

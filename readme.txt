# campaign-automation
Node program that can help with the automation of campaign management

This program is designed to be a time saving application that can process Excel files received from Facebook into a desired format.
Installing it properly might be a little bit challenging, but after installation, usage should be easy enough.

## Installing  
Before Ex2Ex is ready for usage, you need to install some dependencies to make the program run properly.
Below is a step-by-step guide to get you to the process. 

### Opening the terminal
Ex2Ex will be installed and used through the terminal. 
To get to the terminal, open your Spotlight Search with CMD + Space.
Then, type in the word 'Terminal' in the search bar. After pressing return, the terminal should open.

### Homebrew
Now that we've got the terminal open, let's install our first dependency, named Homebrew.
To install Homebrew copy paste the following in your terminal:

    /usr/bin/ruby -e "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/master/install)"
    
You will be asked to confirm you installation by pressing return.
Afterwards, the terminal will ask you to give a password. This is the password of your computer user account.
It might look like the terminal is not registering your password because it's not displaying the characters you type.
Don't worry though, because it certainly is!
After correctly filling in your password and pressing return, Homebrew will install.

### Node.js & NPM
Now that we've got Homebrew up and running, we're halfway there. 
We will use brew to simplify the installing of our next dependencies: are Node.js and NPM. 
To install them, simply type the following command in the terminal:

brew install node

This might take a few minutes. 
If you want to check if your installation worked properly type in:

    node -v 

and / or

    npm -v

These commands should display a version number if everything went right.

### Navigating to the folder
Using the terminal, navigate to the folder where you've installed ex2ex with the following command:

    cd /examplePath/examplePath/examplePath

The easiest way to find a folders path is to drag and drop a folder into the terminal, just like you would drag and dorp a file into outlook.com to add it as an attachment for your email. 
After navigating to the folder, the titlebar of your terminal should change to 'ex2ex' or whatever you named the folder.
If you want to make sure that you are in the right folder, type the following into the terminal:

    pwd

after pressing return, this should display your current directory.

### Installing the NPM packages

Once you're in the right folder, all you need to do is type the following in the terminal:

    npm install

This should automatically install all the NPM packages you need.
It might take a while, but afterwards your ex2ex is ready for usage :) 


## usage
Ex2Ex is run through the terminal. If everything is installed correctly, using it should be very easy.
The commands in the terminal should look as follows:

    node /users/example/ex2ex/src/ex2ex.js /users/example/documents/oldExcel.xlsx /users/example/documents/newExcel.xlsx

The command consists of four parts. The first is the 'node' command. The second is the location of the ex2ex.js file. The third is the Excel file you wish to process. The fourth is the name and location of the new Excel file after processing.
The path of the ex2ex.js file and the old Excel file are easily obtained by simply dragging and dropping the file into the terminal, just like you would when attaching a file to an email.
Always remember to add the .xlsx extension to the new Excel file in the command line, or else it won't render as a readable Excel sheet


## warnings 
Please keep in mind that this program is specifically designed to work with Facebok's 'Weekly Ad Report'. 
Usage with Excel sheets in different formats will lead to errors and unexpected behavior.

As of the writing of this readMe, there is a problem with the new Excel files which sometimes cause a warning to pop up when you wish to open the new file with Excel. 
(Something like: 'We found a problem with some content in 'newExcel.xlsx'. Do you want us to try to recover as much as we can? If you trust the source of this workbook, click Yes.')
Please click 'yes' when prompted to and the file should open normally.

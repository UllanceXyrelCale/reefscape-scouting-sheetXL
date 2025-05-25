# Reefscape Scouting Sheet Program
An IDE-run program that sorts the data of a robot during a match in the Reefscape Season. It calculates the total and average points and sort them in an excel file.

## How It's Made:

### Tech Used:

Language: Python

Library: openpyxl

IDE: PyCharm


### Project Structure:
All the modules are located in a single directory (because I didn't know what I was doing). The main.py module controls all the functions in the program such as, converting the amount of points, caluclating averages, and sorting the datas. The module controller.py allows the usre to take notes of the robot's performance and add them inside the excel file. 

### How it Works
The user changes the value of variables in a controller modules. This is how the program stores the data of a robot during a match. Running this modules allows the program to transfer the data inside the excel file together with different raw datas.

Running the main.py allows the program to make all the necessarry calculations on the given data. The program uses functions that multiply the amount of game objects scored by the robot into their respective point values and added together. 


## Lessons Learned:

I learned how to use the os module in order to check if a file is opened. This allowed the program to avoid crashing when ran if the file is open

I learned how to add modules from a library to my own program

I learned to sonewhat properly structure my program

I learned how to add my program from PyCharm to Github

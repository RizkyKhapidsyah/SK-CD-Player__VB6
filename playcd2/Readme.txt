                  ---------------------------------
	          CDPlyr Sample Project Readme File
		              January l997
		  ---------------------------------
                   (c) Microsoft Corporation, 1997

SUMMARY
=======

CDPLYR.EXE is a self-extracting executable file that contains a Visual
Basic project demonstrating how to use the mciSendString API function
to add multimedia capabilities to your Visual Basic program. The sample
project is an audio CD player.

MORE INFORMATION
================

When you run the self-extracting file, the following files are expanded
into the CD Player Sample Project directory.

 - CDPlyr.vbp
 - CDPlyr.vbw
 - Form1.frm
 - Module1.bas
 - Readme.txt--you are currently reading this file.

The API function mciSendString allows you to send a command string to an
MCI device. You can send commands to detect an audio CD, determine the
number of tracks on the disc, and play the disc at any track or time.
This project shows you how to use this function in Visual Basic.

When you run the program, you must have the AutoPlay feature of the
operating system disabled. Otherwise, unexpected behavior may occur.

You can disable AutoPlay by pressing the SHIFT key when inserting an audio
CD or by completing the following steps.

To Disable AutoPlay
-------------------
1. Double-click the My Computer icon.
2. On the View menu, click Options, and then click the File Types tab.
3. Click the AudioCD type, and then click Edit.
4. In the Actions list, click Play, and then click Set Default.

To Run the Project
------------------

These instructions assume you have disabled AutoPlay.

1. Open the project in Visual Basic.

2. On the Run menu, click Start or press the F5 key to start the program.
   Note that all the command buttons are disabled.

3. Put an audio CD into the CD player of your computer.

The programs starts playing the first track in your audio CD.

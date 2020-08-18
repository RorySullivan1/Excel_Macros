Hello!

Thank you for looking at my automatic / fluid VBA table of contents creator. 
The code.txt file will allow you to simply copy and paste the code into your VBE and run the macro accross any .xlsm document available. 
Please see instructions / changelogs below!

-----INSTRUCTIONS------
In order to run the code it is recommended to attach to a quick toolbar.
The remainder of the code is self-explanitory. 

-----Planned Changed-----
The developer button is unused but will be protecting/hiding sheets that are solely for developers within the excel document. 
This will allow a UI based .xlsm to have more client-facing functionality while supporting continuious development.
	- 'Developer Mode'
	- Better UI for Table of Contents
		- Centered
		- Better zoom level
		- Custom Button design?

----VERSION 1.0.1-----
	-Primary Changes:
		-Focus on formatting of the Topsheet layer and correct implementation of the 'Refresh Button'
			-Addition of Refresh subprocedure
		-Process timing reduced to .03 avg seconds
			-Tested on a 100+ sheet workbook --> .20 seconds
			-Formatting for "Autofiting" the description names on the topsheet is now derived from a mathmatical function
				between columns c/d

Hello!

Thank you for looking at my automatic comments recorded. 
This macro will enable you to generate an additional sheet with any marked 'notes' within the xlsm file. 
I believe this to be important for client-facing reports that are scrutinized and need a level of review prior to distribution. 
This will allow review notes to be easily compiled, forming an agenda for the developer to quickly move through requested changes.

-----INSTRUCTIONS------
In order to run the code it is recommended to attach to a quick toolbar.

PLEASE ONLY USE EXCEL'S 'NOTES' FUNCTION TO RECORD COMMENTS, NOTES WILL ONLY BE RECORDED IN VERSION 1.0.1

The remainder of the code is self-explanitory. 

-----Planned Changed-----
	Primary Changes:
		- 'Threaded Comments' are excel 2016's new version of comments and have compatibility issues with VBA's read function.
		the objects; 'CommentsThreaded' & 'CommentThreaded' were not registred with the property '.value' nor '.text' making compiling
		its input text much more difficult. This is a priority of the next changes to the macro.
	Secondary Changes:
		- Notes with an author name such as; "Rory Sullivan:" will be recorded and input into the text column. 
		This will be resolved moving forward.
		- Refresh on new creation of comment.

----VERSION 1.0.1-----
	The comments section comes alive!
		- The macro will add a new sheet and compile all recorded notes 


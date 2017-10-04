# odt2dita
Automated conversion from ODT (and thus Word) to rough DITA

This Python 2 script converts an ODT document to a set of DITA topics plus map. 

No styles except headers are required. 

Python 2.7 is required. A rudimetary GUI and a command line interface are supported. Start the script without parameters to view the GUI. If the GUI fails with a Tix-related message, use the command line.

Command line syntax:

    python odt2dita-oasis.py infile.odt path/to/output/dir

The input file must exist. For the GUI, the output directory must exist. For the command line, if the output directory does not exist, the script attempts to create it.

**EXISTING FILES IN THE OUTPUT DIRECTORY CAN BE OVERWRITTEN.** Nothing is deleted specially, but if there is a file of the same name as one created, it will be gone. Best practice: create a new directory or let the script create one.

The command line has a few possible options, run the script with the -h option to view them. They correspond to checkboxes and entry fields in the GUI. If you need more options and they can be mplemented logically, please ask for them.

The code has been developed by Mikhail Ramendik since 2010 within IBM and is now open source. 

**This code is prototype/proof of concept quality**. Try it at your own risk. 

The result is "rough" DITA. You should improve it manually to get proper topic-based documentation with content-based markup.

All headers become topics. By default all topics are concepts.

To make a header into a task topic, add \[t\] to the header title (for example: *Writing a letter \[t\]*). To make a header into a reference topic, add \[r\] to the header title (for example: *List of letters \[r\]*).

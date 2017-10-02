# odt2dita
Automated conversion from ODT (and thus Word) to rough DITA

This Python 2 script converts an ODT document to a set of DITA topics plus map. 

No styles except headers are required. A rudimentary UI is present. Just run the script with Python 2.7 to use it. Python 3 is not supported.

The code has been developed by Mikhail Ramendik since 2010 within IBM and is now open source. 

*This code is prototype/proof of concept quality*. Try it at your own risk. 

The result is "rough" DITA. You should improve it manually to get proper topic-based documentation with content-based markup.

All headers become topics. By default all topics are concepts.

To make a header into a task topic, add \[t\] to the header title (for example: *Writing a letter \[t\]*). To make a header into a reference topic, add \[r\] to the header title (for example: *List of letters \[r\]*).

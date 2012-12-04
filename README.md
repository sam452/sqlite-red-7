Bixby applet
=================================

Goal
---------------------------------

This applet is designed to be used by Adobe Indesign users to automate building sets of finished Indesign files populated by
a CSV data file using Applescript. It uses Ruby, ActiveRecord, sqlite3 and Applescript, all of which are found on OS X machines.

Features
--------------------
Beginning with a CSV data file, and a set of Indesign files properly set up, it will begin by dropping a specific squlite3
data file and rebuilding it with the new data file. It uses the first row to determine the column names. From there it 
will begin creating new Indesign files based on each row of the data. It will then look for the first Indesign file 
named in a particular way. This implementation parses the last two or three of the Indesign
file name and begins work on it. The order is built in as a variable. It searches for instances of each field name in the
data and swaps out the field/key in the file with the value for that row. When it finishes each key it produces a PDF
of the Indesign file.

Additionally, it uses pdfjoin to combine the set of pdfs created for each row in the data in the same predetermined order
for presentation.

###Dependencies
This file requires Adobe Indesign CS5 and the stock software found on OS X Leopard, 10.5. It also requires pdfjoin for the
PDF package capability.

###Resources
The applet depends on handler libraries supplied by Bruce Phillips, Craig Williams and many others found at MacScripter.com.

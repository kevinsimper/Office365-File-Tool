Office365-File-Tool
===================

A tool for preparing files to be uploaded to Office 365

Features:
* Make a report over conflicts
* Removes restricted symbols
* Puts restricted file types in a zip file

Made with Python 2.7 and WxWidgets

The report defined 3 kinds of errors:
* Length Error
  The final URL when the file is uploaded to Office 365 must no be longer than 256 character,
  and if there is spaces in the filename they are counted as 3 characters.
* Symbol Error
  There is a lot of symbols that is not allowed like nearly all Unicode characters,
  but there is a dozen of symbols that Windows OS allow, but is not allow on Office 365
* Extension Error
  There is a number of filetypes there is not allowed on Office 365

# Jack-the-Ripper
Jack the Ripper -- Parsing script
CRiTs indicator auto parser (Jack the Ripper)
Using the raw text from the wiki EDIT tab box, copy/paste into your "ToBeParsed.txt" file.
The parser will then go line by line using regex to ignore or pull selected information from lines.
Features:
  Automated indicator parsing
  File extension identification and seperation from attached MD5 hash
  Seperate sender names from email addresses
  Removal of excess characters
  KillSwitch to stop all parsing at a specified point, saving cycles
  Results file unique value sweep (CopyKiller)
    CopyKiller exception to allow multiple file names in case of different MD5s

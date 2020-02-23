# Email Address Tester

An application tha tests email addresses to see if they actually exists.

### Prerequisites

* CSV file containing email addresses. Make sure the addresses are all in column A in Excel.
* Make sure the filename is 'emails.csv'.
* Need to place the CSV file in \bin\Debug if you are running this in Visual Studio 2017.

### What is this?

This is not to test email address syntax.
This application finds out a mx record of given email address and send SMTP commands including 'rcpt to:'.
Based on the response after issuing 'rcpt to:' command to corresponding exchange server, it determines whether an address is valid or not.

### Caveat

* This application may not work properly from residential network as telnetting to port 25 is likely blocked by ISP.* 
* If their exchange policy is set to 'catch-all', there is no way know if an address exists and the application thinks it's valid.

### Future Improvements

* Take input from command line to specify where the program needs to load CSV file from.
* Take input from command line to specify where an output report gets generated.

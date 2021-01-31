# Get-NetworkAdapterSettings

This will use PS Remoting to get a machine's network adapter settings and output to multiple formats.  Input is an array of computer names.  If no computer name is specified, a list of all computers in the domain is created.  The results for the network settings or a DataTable version of the results can be returned so the machine originating PS Remoting can further process them.  File output for all remote machines is to a single file for each filetype on the machine originating PS Remoting and can be to an HTML, CSV, or list-formated text file.  The HTML or CSV file can be opened in Excel for further processing.  I filtered out the network adapter types 'Miniport', 'ISATAP', and 'Debug' and they greatly increased the number of results and I didn't have a need to look at them.  You can reinclude them by commenting out those three lines.  You can also increase the throttle limit on Invoke-Command as it currently is set to sequencially connect to each machine.

![Output Sample](https://github.com/donhess321/Get-NetworkAdapterSettings/blob/main/output_sample.png)

This is a reposting from my Microsoft Technet Gallery.

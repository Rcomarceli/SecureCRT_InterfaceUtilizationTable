# SecureCRT_InterfaceUtilizationTable
This program is currently confirmed to be working on version 6.5.4 of SecureCRT, a SSH client that allows for scripting via VbScript.
crt.screen commands below interact with the SecureCRT client, mostly telling it to input keys or wait for certain output before continuing.
Hostnames and Hostname structures have been substituted for security purposes.

The objective of the program was to:
- Mostly automate a NOC task where technicians needed to gather transmit and receive utilization percentages for certain hosts and their interfaces for a daily report
- Be compatiable with *every* NOC engineer's environment without the need to install a coding language or program, as all users lacked install permissions.
For compatiability, I coded this to work with SecureCRT and output the data in a formatted table via printf in a linux environment, as this was the common environment for every NOC engineer. 
For implementation, the program logs into devices, grabs transmit and receive utilizations for certain interfaces, processes that data into percentages, and displays that data in a readable table 
More specifically, it ssh's into a list of defined hostname and interface pairs, shows interface output and grabs the data via Regex, processes the data, and outputs a printf command that when entered on a Linux server, displays a formatted table of data that looks like below:


Site1     TB Router 1     TB Router 1     TB Router 2     TB Router 2     Edge Router 1 
Int:      Te1/0/5         Te1/0/3         Te1/0/5         Te1/0/3         Te0/1/0 
Tx:       0.4%            13.7%           3.9%            11.8%           11.8% 
Rx:       0.4%            11%             3.5%            0.4%            3.5%
 
Site2     TB Router 1     TB Router 1     TB Router 2     TB Router 2     Edge Router 1 
Int:      Te1/0/2         Te1/0/3         Te1/0/2         Te1/0/3         Te0/1/0 
Tx:       2%              0.4%            0.4%            1.2%            5.9% 
Rx:       1.2%            0.4%            0.4%            2.4%            2.4%
 
Site3     TB Router 1 
Int:      Gi0/0/5 
Tx:       0.8% 
Rx:       1.2%

This specific table format was chosen to mimic the Onenote table these data values would later be manually input to.

# README #

This application is meant to help quickly clear out old profiles on high-traffic machines that cannot be made auto-logon. Helpful for speeding up computer login times, drive encrypt/decrypt times, and general device maintenance. 

### Quick Summary ###

* Program written in Visual Studio 2015 w/ Resharper
* Utilizes WMI query and the System.Management namespace to have the system delete old profiles. 
* Current version: 1
 
 ### Commands ###
 * /device:<deviceID> : Target Windows PC. Usage: /device:TESTCOMPUTER
 * /days:<x> : Specifiy minimum age of profiles to delete. Usage: /days:14
 * /stopwatch : Tracks the speed of profile deletion and the entirety of the process, then prints out performance information.
 * /quiet || /q : Silences output of the program. Very little information will be printed to the console.
 
 By entering no commands, 

### Required External Components ###
* No external requirements!

### How do I get set up? ###
* Compile the program and run the executable
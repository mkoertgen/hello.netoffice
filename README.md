# Hello.NetOffice

A sample for a self-registrating Office Addin using elevation.

## What is it?

[NetOffice](https://github.com/NetOfficeFw/NetOffice) is a great framework for writing Microsoft Office AddIns.
When deploying an Addin you need to register it for COM which requires administrative privileges. 

The usual approach for this is an installer. An alternative is to use ClickOnce with VSTO Runtime (Note that you don't need VSTO with NetOffice). 

What if your AddIn could just register itself? This is exactly what the sample does. The AddIn is wrapped together with a simple program
registering for COM, then starting Office (here: PowerPoint).  



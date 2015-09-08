Mingle Membership Request Automation 
====================================
System: Windows
Applications: Strawberry Perl, Mingle Thoughtworks Agile
Email Application: LotusNotes

Usage: membership_request.pl 
Purpose: To automate manual clicking of emails to approve existing Mingle users who requires additional access to projects or programs.

* Window Task Schedueler execute "membership_request.pl" every 5 minutes
* requires Win32::OLE to access local LotusNotes database.nfs to query LotusNotes emails
* create/output data to a .csv file for every users who have requested membership access


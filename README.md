Mingle Membership Request Automation (LotusNotes, os: Windows)
====================================
Requirements:
Windows OS 
LotusNotes Client 8.5.3
Strawberry Perl  http://strawberryperl.com/
Curl.exe

Setup CURL on Windows 7 64-bit
Setup CURL:
1.	Download and unzip 64-bit cURL with SSL: http://curl.download.nextag.com/download/curl-7.21.7-win64-ssl-sspi.zip 
2.	Copy the curl.exe file into your Windows PATH folder.
3.	Download and install the Visual Studio 2010 C++ Runtime Redistributable 64 bit here: http://www.microsoft.com/download/en/details.aspx?id=13523
4.	Download the latest bundle of Certficate Authority Public Keys from mozilla.org: http://curl.haxx.se/ca/cacert.pem
5.	Rename this file from cacert.pem to curl-ca-bundle.crt.
6.	Move this file into your Windows PATH folder
7.	Download msvcr100.dll file if this file is missing

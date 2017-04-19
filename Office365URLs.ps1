<# 
Office 365 URL generator based on Office 365 XML feed:

https://support.content.office.net/en-us/static/O365IPAddresses.xml

THIS CODE AND ANY ASSOCIATED INFORMATION ARE PROVIDED “AS IS” WITHOUT WARRANTY 
OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE 
IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR
PURPOSE. THE ENTIRE RISK OF USE, INABILITY TO USE, OR RESULTS FROM THE USE OF 
THIS CODE REMAINS WITH THE USER.

Author:		Sarah Lean
Credit to Aaran Guilmette for the original script from which this is modified. 
 
Find the most updated version at:

Change Log: 
V1.00, 06/12/2016 - Initial version

#>

<#
.SYNOPSIS
I have taken the script that was written by Aaron Guilmette (https://gallery.technet.microsoft.com/Office-365-Proxy-Pac-60fb28f7)
and modified it.  This script queries the latest Office 365 URLs website (https://support.content.office.net/en-us/static/O365IPAddresses.xml)
and pulls the information there to a text file.  The script appends the correct start/end to each URL so that it can be utilised within a PAC file. 
The script sorts and removes any duplicates from the list generated from the XML, and also stripes out any references to Facebook, Google, Dropbox
and Hockey. 

The generated text file can be used to help construct a PAC file for your environment. 


.PARAMETER Products
Use the Products parameter to specify which products will be configured in the
PAC. The full list of products keywords that can be used:
	'O365' - Office 365 Portal
	'LYO' - Skype for Business (formerly Lync Online)
	'Planner' - Planner
	'ProPlus' - Office 365 ProPlus
	'OneNote' - OneNote
	'WAC' - SharePoint WebApps
	'Yammer' - Yammer
	'EXO' - Exchange online
	'Identity' - Office 365 Identity
	'SPO' - SharePoint Online
	'RCA' - Remote Connectivity Analyzer
	'Sway' - Sway
	'OfficeMobile' - Office Mobile Apps
	'Office365Video' - Office 365 Video
	'CRLs' - Certificate Revocation Links
	'OfficeiPad' - Office for iPad
	'EOP' - Exchange Online Protection

.PARAMETER temp
The temp parameter specifies the name of the first temporary file. This is where all the data from
the XML is dumped on first pass. 

.PARAMETER temp2
The temp parameter specifies the name of the second temporary file. This is where the URLS are put 
after any references to Facebook, Dropbox, Hockey or Dropbox have been removed. 

.PARAMETER OutputFile
The OutputFile parameter specifies the name of the output PAC file.

.EXAMPLE
.\Office365URLs.ps1 
This will just run the script with all defaults. 

.EXAMPLE
.\Office365URLs.ps1 -OutputFile Proxy.txt
This will run the script but change the name of the final output file. 

.EXAMPLE
.\Office365URLs.ps1 -Products EXO,LYO
This will run the script and only include entries related to Exhange Online and Skype for Business (formerly Lync Online).

#>
 
[CmdletBinding()]
Param(
[Parameter(Mandatory=$false,HelpMessage='TempFile')]
		[string]$temp = "temp.txt",
    [Parameter(Mandatory=$false,HelpMessage='Temp2File')]
		[string]$temp2 = "temp2.txt",
    [Parameter(Mandatory=$false,HelpMessage='OutputFile')]
		[string]$OutputFile = "Office365URLs.txt",
	[ValidateSet("O365","LYO","Planner","ProPlus","OneNote","WAC","Yammer","EXO","Identity","SPO","RCA","Sway","OfficeMobile","Office365Video","CRLs","OfficeiPad","EOP")]
		[array]$Products = ('O365','LYO','Planner','ProPlus','OneNote','WAC','Yammer','EXO','Identity','SPO','RCA','Sway','OfficeMobile','Office365Video','CRLs','OfficeiPad','EOP')
	)

Write-Host "The PAC file will be generated for the following products:"
Write-Host $Products

[regex]$ProductsRegEx = ‘(?i)^(‘ + (($Products |foreach {[regex]::escape($_)}) –join “|”) + ‘)$’

	{
	$O365URL = "https://support.content.office.net/en-us/static/O365IPAddresses.xml"
	Write-Host -ForegroundColor Yellow "Downloading latest Office 365 XML data..."
	[xml]$O365URLData = (New-Object System.Net.WebClient).DownloadString($O365URL)
	}

$SelectedProducts = $O365URLData.SelectNodes("//product") | ? { $_.Name -match $ProductsRegEx }

#Remove any files from previous attempts
If (Test-Path $temp2) { Remove-Item -Force $temp2 }
If (Test-Path $temp) { Remove-Item -Force $temp }
If (Test-Path $OutputFile) { Remove-Item -Force $OutputFile }
 

$ProxyURLData = @()


$SelectedProducts.AddressList | ? { $_.Type -eq "URL" } | % { $address = $_; foreach ($a in $address) { $ProxyURLData += $a.address } }

# Build Proxy List
Foreach ($url in $ProxyURLData)
	{
	Write-Host $url

	If ($url -match "\*")
		{
		Add-Content $temp "shExpMatch(host, ""$URL"")||"
		}
	Else
		{
		Add-Content $temp "dnsDomainIs(host, ""$URL"")||"
		}
	}


#Remove duplicate records, unwanted URLs and outputs to final file
gc $temp | select-string -pattern '(Facebook)|(youtube)|(hockey)|(dropbox)|(google)' -notmatch > $temp2
gc $temp2 | Sort-Object -Unique > $OutputFile

#Removes temp file
Remove-Item -Force $temp
Remove-Item -Force $temp2



Try {
	Test-Path $Outputfile -ErrorAction SilentlyContinue > $null
	Write-Host -ForegroundColor Yellow "Done! Office365 URLs have been outputted to file $OutputFile."
	}
Catch {
	Write-Host -ForegroundColor Red "The file was not created."
	}
Finally { }
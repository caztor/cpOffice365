<# 
Automatic CheckPoint network object creator for Office 365 Address space
based on the official Office 365 XML feed:

https://support.content.office.net/en-us/static/O365IPAddresses.xml

THIS CODE AND ANY ASSOCIATED INFORMATION ARE PROVIDED “AS IS” WITHOUT WARRANTY 
OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE 
IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR
PURPOSE. THE ENTIRE RISK OF USE, INABILITY TO USE, OR RESULTS FROM THE USE OF 
THIS CODE REMAINS WITH THE USER.

Author:		Theis Andersen Samsig
          tas@dubex.dk

Inspired by the Automatic PAC file generator script by Aaron Guilmette found at:
https://gallery.technet.microsoft.com/Office-365-Proxy-Pac-60fb28f7
#>

<#
.SYNOPSIS
Automatically imports 'Office 365 URLs and IP address ranges' XML feed into CheckPoint 
utilising psCheckPoint powershell module.

.PARAMETER XMLFile
Use specified XML import file instead of downloading from support site.

.PARAMETER Products
Use the Products parameter to specify which products will be imported.
The full list of products keywords that can be used:
	- 'O365' - Office 365 Portal and Shared
	- 'LYO' - Skype for Business (formerly Lync Online)
	- 'Planner' - Planner
	- 'ProPlus' - Office 365 ProPlus
	- 'OneNote' - OneNote
	- 'WAC' - SharePoint WebApps
	- 'Yammer' - Yammer
	- 'EXO' - Exchange online
	- 'Identity' - Office 365 Identity
	- 'SPO' - SharePoint Online
	- 'RCA' - Remote Connectivity Analyzer
	- 'Sway' - Sway
	- 'OfficeMobile' - Office Mobile Apps
	- 'Office365Video' - Office 365 Video
	- 'CRLs' - Certificate Revocation Links
	- 'OfficeiPad' - Office for iPad
	- 'EOP' - Exchange Online Protection
	- 'EX-Fed' - Exchange Federation (?)
  - 'Teams' - Microsoft Teams

.EXAMPLE


.LINK


.NOTES
2018-02-07  First working version released

.TODO
Can't remove from just a single product group as it removes the object from all other groups
Not adding multiple tags to same network object
#>
 
[CmdletBinding()]
Param(
	[ValidateSet("O365","LYO","Planner","ProPlus","OneNote","WAC","Yammer","EXO","Identity","SPO","RCA","Sway","OfficeMobile","Office365Video","CRLs","OfficeiPad","EOP","EX-Fed","Teams")]
		[array]$Products = ('O365','LYO','Planner','ProPlus','OneNote','WAC','Yammer','EXO','Identity','SPO','RCA','Sway','OfficeMobile','Office365Video','CRLs','OfficeiPad','EOP','EX-Fed','Teams'),

	[Parameter(Mandatory=$false,HelpMessage='Default tag')]
		[string]$DefaultTag = "Office365",
	
	[Parameter(Mandatory=$false)]
		[string]$XMLFile = "O365IPAddresses.xml",

	[Parameter(Mandatory=$false)]
		[string]$CommentPrefix = "MS Office365",

	[Parameter(Mandatory=$false)]
		[string]$NetworkPrefix = "Net_",

	[Parameter(Mandatory=$false)]
		[string]$GroupPrefix = "O365",

	[Parameter(Mandatory=$false,HelpMessage='OutputFile')]
		[string]$OutputFile = "Office365CP.txt",

	[Parameter(Mandatory=$false,HelpMessage='Office 365 XML file')]
		[string]$O365URL = "https://support.content.office.net/en-us/static/O365IPAddresses.xml"
	)

Function cidr
{
	[CmdLetBinding()]
	Param (
		[Parameter(Mandatory = $True, Position = 0, ValueFromPipeline = $True)]
		[Alias("Length")]
		[ValidateRange(0, 32)]
		$MaskLength
	)
	Process
	{
		Return LongToDotted ([Convert]::ToUInt32($(("1" * $MaskLength).PadRight(32, "0")), 2))
	}
}

Function LongToDotted
{
	[CmdLetBinding()]
	Param (
		[Parameter(Mandatory = $True, Position = 0, ValueFromPipeline = $True)]
		[String]$IPAddress
	)
	Process
	{
		Switch -RegEx ($IPAddress)
		{
			"([01]{8}\.){3}[01]{8}" {
				Return [String]::Join('.', $($IPAddress.Split('.') | ForEach-Object { [Convert]::ToUInt32($_, 2) }))
			}
			"\d" {
				$IPAddress = [UInt32]$IPAddress
				$DottedIP = $(For ($i = 3; $i -gt -1; $i--)
					{
						$Remainder = $IPAddress % [Math]::Pow(256, $i)
						($IPAddress - $Remainder) / [Math]::Pow(256, $i)
						$IPAddress = $Remainder
					})
				Return [String]::Join('.', $DottedIP)
			}
			default
			{
				
			}
		}
	}
}

$username = $env:UserName.ToUpper()
$updateRequired = 0

Write-Host "The CP file will be generated for the following products:"
Write-Host $Products

[regex]$ProductsRegEx = ‘(?i)^(‘ + (($Products |foreach {[regex]::escape($_)}) –join “|”) + ‘)$’

If (Test-Path $XMLFile) {
    [xml]$O365URLData = Get-Content $XMLFile
    $fileUpdated = [datetime]($O365URLData.products.updated)
}

	Write-Host -ForegroundColor Yellow "Fetching latest online Office 365 XML data ..."
	[xml]$O365URLData = (New-Object System.Net.WebClient).DownloadString($O365URL)

$onlineUpdated = [datetime]($O365URLData.products.updated)
$formatUpdated = "{0:yyyyMMdd}" -f $onlineUpdated

if($fileUpdated -lt $onlineUpdated) {
    $updateRequired = 1
}

if($updateRequired -or (-Not (Test-Path $XMLFile))) {
    Write-Host -ForegroundColor Red "Backup XML file not found or outdated!"
    if(Test-Path $XMLFile) { Rename-Item -Force -Path $XMLFile -NewName ($XMLFile -replace "\.","_$("{0:yyyyMMdd}" -f $fileUpdated).") }
    Invoke-WebRequest $O365URL -OutFile $XMLfile;
}

$SelectedProducts = $O365URLData.SelectNodes("//product") | ? { $_.Name -match $ProductsRegEx }

$IPData = @()
$AlwaysProxyURLMatches = @()

foreach ($Product in $SelectedProducts) {

foreach ($AddressList in $Product.addresslist) {

	$AddressList | ? { $_.Type -eq "IPv4" } | % { foreach ($a in $_.address) { $IPData += @{subnet=$a; product=$Product.name} } }

}

}

#echo $IPData | %{echo $_;}

# Run the update if required
#$ProxyURLData = $ProxyURLData | Sort -Unique

if($updateRequired = 1) {

foreach ($Product in $SelectedProducts) {
    if(-Not (Get-CheckPointGroup -Name "$($GroupPrefix)_$($Product.name)" -ErrorAction SilentlyContinue)) {
        Write-Host -ForegroundColor Yellow "Creating new group $($GroupPrefix)_$($Product.name)"
        New-CheckPointGroup -Name "$($GroupPrefix)_$($Product.name)" -Comments "$($CommentPrefix): $($Network.product) - $username $($formatUpdated)"
    }
    Write-Host -ForegroundColor Green "Removing members from Group $($GroupPrefix)_$($Product.name)"
    $groupMembers = Get-CheckPointGroup "$($GroupPrefix)_$($Product.name)"
    $groupMembers.Members | where { $_.Name -like "$($NetworkPrefix)*" } | % { Write-Host -ForegroundColor Yellow "Removing $_"; Get-CheckPointNetwork -Name $_ | % { Write-Host -ForegroundColor Yellow " ... from group $($_.Groups)"; Set-CheckPointNetwork -Name $_ -GroupAction Remove -Groups $_.Groups } ; Remove-CheckPointNetwork -Name $_ }
}

Write-Host -ForegroundColor Yellow "Adding members to groups ..."
Foreach ($Network in $IPData)
	{

	If ($Network.subnet -match "\b(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\b")
		{
				$CIDR = $Network.subnet.Split("/")[1]
				$IPAddr = $Network.subnet.Split("/")[0]
                Write-Host -ForegroundColor Green "Adding $($NetworkPrefix)$($Network.subnet)"

                if((Get-CheckPointNetwork -Name "$($NetworkPrefix)$($Network.subnet)" -ErrorAction SilentlyContinue)) {
                    Set-CheckPointNetwork -Name "$($NetworkPrefix)$($Network.subnet)" -GroupAction Add -Groups "$($GroupPrefix)_$($Network.product)"
                    Write-Host -ForegroundColor Green " ... (existing) to group $($GroupPrefix)_$($Network.product)"
                } else {
                    New-CheckPointNetwork -Name "$($NetworkPrefix)$($Network.subnet)" -Subnet $($IPAddr) -MaskLength $($CIDR) -tags @($DefaultTag,$($Network.product)) -Comments "$($CommentPrefix): $($Network.product) - $username $($formatUpdated)" -Groups "$($GroupPrefix)_$($Network.product)"
                    Write-Host -ForegroundColor Green "Creating new network $($NetworkPrefix)$($Network.subnet)"
                }

		}
}

}

If (Test-Path $OutputFile) { Remove-Item -Force $OutputFile }

Try {
	Test-Path $OutputFile -ErrorAction SilentlyContinue > $null
	Write-Host -ForegroundColor Yellow "Done! CP file is $($OutputFile)."
	}
Catch {
	Write-Host -ForegroundColor Red "CP file not created."
	}
Finally { }

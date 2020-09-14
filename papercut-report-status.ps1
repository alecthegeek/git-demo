#requires -version 2
<#
.SYNOPSIS
Report on PaperCut MF/NG Status
.DESCRIPTION
Displays a dump of PaperCut related information
.PARAMETER <None>
.INPUTS
None
.OUTPUTS
Outputs Demo Report
.NOTES
Version:        0.1
Author:         Alec Clews
Creation Date:  2020/September
Purpose/Change: Initial script development

.EXAMPLE
<Example goes here. Repeat this attribute for more than one example>:
#>


Add-Type -Path "$env:USERPROFILE\.nuget\packages\kveer.xmlrpc\1.1.1\lib\netstandard2.0\Kveer.XmlRPC.dll"
Add-Type -Path "$PWD\ServerCommandProxy\bin\Release\netstandard2.0\ServerCommandProxy.dll"

# If not localhost then this address will needs to be listed as allowed in PaperCut
$papercuthost = "localhost"

$auth = "token" # Value defined in advanced config property "auth.webservices.auth-token". Should be random

# Proxy object to call PaperCut Server API
$s = New-Object PaperCut.ServerCommandProxy($papercuthost, 9191, $auth)
# Find the value of the Auth key in the PaperCut MF/NG config database
$authString = $s.GetConfigValue("health.api.key")

# Set up the http header for the health API 
$headers = @{'Authorization' = $authString}

$uri = [Uri]"http://localhost:9191/api/health"

# Generate the report
&{
  # Get a list of the processes running 
  Get-Service -DisplayName  *PaperCut* | Select-Object -Property name,status |
            ConvertTo-Html -Fragment -PreContent '<h2>PaperCut Processes</h2>'

  $rsp = Invoke-RestMethod -Uri $uri -Method Get -Headers $headers

  $rsp.applicationServer.systemInfo |
  ConvertTo-Html -As List -Fragment -PreContent '<h2>PaperCut Services Info</h2>'

  Write-Output "<p>Total Printers = $($rsp.printers.count)</p>"
  Write-Output "<p>Total Devices = $($rsp.devices.count)</p>"
 } | Out-File -FilePath .\report1.html

 Invoke-Expression .\report1.html

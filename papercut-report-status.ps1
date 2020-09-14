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

Get-Service -DisplayName *PaperCut* | Format-Table -Property  Name, Status, StartType

$epsonErrCount = (Select-String -Path c:\tmp\epsonOPS.log -Pattern error | Measure-Object).count

if ($epsonErrCount -gt 0) {
    write-output "Epson plugin has reported $epsonErrCount errors"
}

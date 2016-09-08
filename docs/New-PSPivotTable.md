---
external help file: PSPivotTable-help.xml
online version: 
schema: 2.0.0
---

# New-PSPivotTable
## SYNOPSIS
Create a pivot table in the PowerShell console from PowerShell commands.

## SYNTAX

### Property (Default)
```
New-PSPivotTable [-Data] <Object> [-yLabel <String>] -yProperty <String> -xLabel <String> [-xProperty <String>]
```

### Count
```
New-PSPivotTable [-Data] <Object> [-yLabel <String>] -yProperty <String> -xLabel <String> [-Count] [-Sort <String>] [-SortKey <String>]
```

### Sum
```
New-PSPivotTable [-Data] <Object> [-yLabel <String>] -yProperty <String> -xLabel <String> [-Sum <String>] [-Format <String>] [-Round <Int32>] [-Sort <String>] [-SortKey <String>]
```

## DESCRIPTION
This command takes the result of a PowerShell expression and creates a pivot table object. You can use this object to analyze data patterns.
For example, you could get a directory listing and then prepare a table showing the size of different file extensions for each folder.

## EXAMPLES

### -------------------------- EXAMPLE 1 --------------------------
```
PS C:\> $svc="Lanmanserver","Wuauserv","DNS","ADWS"
PS C:\> $computers="chi-dc01","chi-dc02","chi-dc04"
PS C:\> $data = Get-Service -name $svc -ComputerName $computers
PS C:\> New-PSPivotTable $data -ylabel Computername -yProperty Machinename -xlabel Name -xproperty Status | format-table -autosize

Computername    ADWS     DNS Lanmanserver Wuauserv
------------    ----     --- ------------ --------
chi-dc01     Running Running      Running  Running
chi-dc02     Running Stopped      Running  Running
chi-dc04     Running Running      Running  Stopped
```

Create a table that shows the status of each service on each computer.
The yLabel parameter renames the property so that instead of Machinename it shows Computername. 
The xLabel is the property name to analyze, in this case the service name.
The xProperty value of each service becomes the table value.

### -------------------------- EXAMPLE 2 --------------------------
```
PS C:\> $files = dir c:\scripts -include *.ps1,*.txt,*.zip,*.bat -recurse
PS C:\> New-PSPivotTable $files -yProperty Parent -xLabel Extension -count | Export-CSV -path c:\work\scriptdir.csv -notypeinformation
```
Create a report that shows the count of each file type in each top level directory and export to a CSV file.


### -------------------------- EXAMPLE 3 --------------------------
```
PS C:\> $files = (dir -path c:\scripts -file).Where({$_.extension -match "ps1|txt|zip|bat|xml"})
PS C:\> New-PSPivotTable $files -yProperty Directory -xlabel Extension -Sum Length -round 2 -format kb | format-table -auto 

Directory      .TXT   .PS1      .XML   .ZIP  .BAT .PS1XML
---------      ----   ----      ----   ----  ---- -------
C:\scripts 30422.99 9494.2 270941.55 159.71 26.62  139.57
```

Analyse files by extension, measuring the total size of each extension. The value is formatted as KB to 2 decimal points.

### -------------------------- EXAMPLE 4 --------------------------
```
PS C:\> New-PSPivotTable $files -yProperty Directory -xLabel Extension -Count -Sort Ascending

Directory : C:\scripts
.ZIP      : 15
.BAT      : 18
.PS1XML   : 25
.XML      : 121
.TXT      : 480
.PS1      : 2288
```

Process the collection of script files and analyze by the count of each file type.
The result is sorted by the count value in ascending order. Note that the actual output would include the period as part of the extension.

### -------------------------- EXAMPLE 5 --------------------------
```
PS C:\> $files = dir c:\scripts -recurse -File | 
Select *, @{Name="Age";Expression={(Get-Date)-$_.LastWriteTime}},
@{Name="Bucket";Expression={
Switch([int]((Get-Date)-$_.LastWriteTime).TotalDays) {
{$_ -gt 365} {'365Plus' ; Break}
{$_ -gt 180 -AND $_ -le 365} {'1Yr' ; Break}
{$_ -gt 90 -AND $_ -le 180} {'6Mo' ; Break}
{$_ -gt 30 -AND $_ -le 90} {'3Mo' ; Break}
{$_ -gt 7 -AND $_ -le 30} { '1Mo'; Break }
{$_ -gt 0 -AND $_ -le 7} { '1Wk' ; Break }
Default { 'Today' }
} 
}}

PS C:\> New-PSPivotTable $files -yProperty Directory -xLabel Bucket -count | Out-GridView -title "File Aging"
```

Get all files and include some aging information based on the last write time. Then create a pivot table on the directory and aging buckets and display results with Out-Gridview.

### -------------------------- EXAMPLE 6 --------------------------
```
PS C:\> New-PSPivotTable -Data (get-eventlog system -newest 1000) -Count -yProperty EntryType -xLabel Source | Out-Gridview -title 'System Sources'
```

Create a pivot table with a Y column of Entry Type and the X axis labels of the different sources based on the 1000 newest system event logs.
The value under each column will be the total count of entries by source. The results are piped to Out-Gridview for viewing and further sorting or filtering.

### -------------------------- EXAMPLE 7 --------------------------
```
PS C:\> $e = get-eventlog system -newest 1000 -entrytype Error
PS C:\> New-PSPivotTable $e -yProperty EntryType -xLabel Source -count -sort Descending

EntryType                             : Error
DCOM                                  : 915
MICROSOFT-WINDOWS-WINDOWSUPDATECLIENT : 40
SERVICE CONTROL MANAGER               : 15
DISK                                  : 11
EVENTLOG                              : 7
SRV                                   : 4
BTHUSB                                : 2
KERBEROS                              : 2
MICROSOFT-WINDOWS-NDIS                : 1
AX88179                               : 1
BUGCHECK                              : 1
MICROSOFT-WINDOWS-HYPER-V-VMSWITCH    : 1
```

Create a pivot table on the error source, sorted by count in descending order.

### -------------------------- EXAMPLE 8 --------------------------
```
PS C:\> $k = Get-Eventlog -source *kernel* -logname System
PS C:\> New-PSPivotTable $k -yProperty EntryType -xLabel Source -count -sort Ascending -SortKey Name

EntryType                                : Information
MICROSOFT-WINDOWS-KERNEL-BOOT            : 1120
MICROSOFT-WINDOWS-KERNEL-GENERAL         : 2051
MICROSOFT-WINDOWS-KERNEL-PNP             : 0
MICROSOFT-WINDOWS-KERNEL-POWER           : 168
MICROSOFT-WINDOWS-KERNEL-PROCESSOR-POWER : 604

EntryType                                : Warning
MICROSOFT-WINDOWS-KERNEL-BOOT            : 0
MICROSOFT-WINDOWS-KERNEL-GENERAL         : 0
MICROSOFT-WINDOWS-KERNEL-PNP             : 784
MICROSOFT-WINDOWS-KERNEL-POWER           : 0
MICROSOFT-WINDOWS-KERNEL-PROCESSOR-POWER : 1816

EntryType                                : 0
MICROSOFT-WINDOWS-KERNEL-BOOT            : 0
MICROSOFT-WINDOWS-KERNEL-GENERAL         : 0
MICROSOFT-WINDOWS-KERNEL-PNP             : 0
MICROSOFT-WINDOWS-KERNEL-POWER           : 24
MICROSOFT-WINDOWS-KERNEL-PROCESSOR-POWER : 0

EntryType                                : Error
MICROSOFT-WINDOWS-KERNEL-BOOT            : 0
MICROSOFT-WINDOWS-KERNEL-GENERAL         : 9
MICROSOFT-WINDOWS-KERNEL-PNP             : 0
MICROSOFT-WINDOWS-KERNEL-POWER           : 0
MICROSOFT-WINDOWS-KERNEL-PROCESSOR-POWER : 0
```
Create a variable of all entries where the source includes 'kernel' in the name. Then create a pivot table for each entry type showing the count of each source. The results are sorted by the source name.

## PARAMETERS

### -Data
This is the collection of data object to analyze. You must enter a parameter value. See help examples.

```yaml
Type: Object
Parameter Sets: (All)
Aliases: 

Required: True
Position: 1
Default value: 
Accept pipeline input: False
Accept wildcard characters: False
```

### -yLabel
This is an alternative value for the the "Y-Axis". If you don't specify a value then the yProperty value will be used. Use this parameter when you want to rename a property value such as Machinename to Computername.

```yaml
Type: String
Parameter Sets: (All)
Aliases: 

Required: False
Position: Named
Default value: 
Accept pipeline input: False
Accept wildcard characters: False
```

### -yProperty
The property name to pivot on. This is the "Y-Axis" of the pivot table.

```yaml
Type: String
Parameter Sets: (All)
Aliases: 

Required: True
Position: Named
Default value: 
Accept pipeline input: False
Accept wildcard characters: False
```

### -xLabel
The property name that you want to pivot on horizontally.
The value of each corresponding object property becomes the label on the "X-Axis". For example, if the Data is a collection of service objects and xLabel is Name, each column will be labeled with the name of a service object, e.g. Alerter or BITS. See help examples.

```yaml
Type: String
Parameter Sets: (All)
Aliases: 

Required: True
Position: Named
Default value: 
Accept pipeline input: False
Accept wildcard characters: False
```

### -xProperty
The property name that you want to analyze for each object. This will be used for calculating table values.
See help examples.

```yaml
Type: String
Parameter Sets: Property
Aliases: 

Required: False
Position: Named
Default value: 
Accept pipeline input: False
Accept wildcard characters: False
```

### -Count
Instead of getting a property for each xLabel value, return a total count of each.

```yaml
Type: SwitchParameter
Parameter Sets: Count
Aliases: 

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -Sum
Instead of getting a property for each xLabel value, return a total sum of each.
The parameter value is the object property to measure.

```yaml
Type: String
Parameter Sets: Sum
Aliases: 

Required: False
Position: Named
Default value: 
Accept pipeline input: False
Accept wildcard characters: False
```

### -Format
If using -Sum the default output is typically bytes, depending on the object. Use KB, MB, GB or TB to reformat the sum accordingly.

```yaml
Type: String
Parameter Sets: Sum
Aliases: 

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Round
Use this value to round a sum, especially if you are formatting it to something like KB.

```yaml
Type: Int32
Parameter Sets: Sum
Aliases: 

Required: False
Position: Named
Default value: 0
Accept pipeline input: False
Accept wildcard characters: False
```

### -Sort
You have the option of sorting results when using -Count or -Sum. You can sort on the value, i.e. count or sum, or on the property name. The default sort option is none but you can specify Ascending or Descending.

```yaml
Type: String
Parameter Sets: Count, Sum
Aliases: 

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SortKey
Specify if you want to sort on the value or property name. The default is Value. This parameter has no effect unless you also use -Sort.

```yaml
Type: String
Parameter Sets: Count, Sum
Aliases: 

Required: False
Position: Named
Default value: Value
Accept pipeline input: False
Accept wildcard characters: False
```

## INPUTS

### None

This command does not accept any pipelined input.

## OUTPUTS

### PSCustomObject

This command writes a custom object to the pipeline.

## NOTES
* NAME:     New-PSPivotTable
* AUTHOR:   Jeffery Hicks (@JeffHicks)
* VERSION:  2.1.3
* LASTEDIT: 8 September 2016

This function was first published and described at http://jdhitsolutions.com/blog/powershell/2434/powershell-pivot-tables/

Learn more about PowerShell: http://jdhitsolutions.com/blog/essential-powershell-resources/

Thanks to kdoblosky for contributing to this module.

## RELATED LINKS
[Measure-Object]()
[Group-Object]()
[Select-Object]()


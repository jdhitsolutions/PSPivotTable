#requires -version 4.0


Function New-PSPivotTable {

<#
.SYNOPSIS
Create a pivot table in the PowerShell console
.DESCRIPTION
This command takes the result of a PowerShell expression and creates a pivot table object. You can use this object to analyze data patterns. 
For example, you could get a directory listing and then prepare a table showing the size of different file extensions for each folder.
.PARAMETER Data
This is the collection of data object to analyze. You must enter a parameter value. See help examples.
.PARAMETER yLabel
This is an alternative value for the the "Y-Axis". If you don't specify a value then the yProperty value will be used. Use this parameter when you want to rename a property value such as Machinename to Computername.
.PARAMETER yProperty
The property name to pivot on. This is the "Y-Axis" of the pivot table.
.PARAMETER xLabel
The property name that you want to pivot on horizontally. The value of each corresponding object property becomes the label on the "X-Axis". For example, if the Data is a collection of service objects and xLabel is Name, each column will be labeled with the name of a service object, e.g. Alerter or BITS. See help examples.
.PARAMETER xProperty
The property name that you want to analyze for each object. This will be used for calculating table values. See help examples.
.PARAMETER Count
Instead of getting a property for each xLabel value, return a total count of each.
.PARAMETER Sum
Instead of getting a property for each xLabel value, return a total sum of each. The parameter value is the object property to measure.
.PARAMETER Format
If using -Sum the default output is typically bytes, depending on the object. Use KB, MB, GB or TB to reformat the sum accordingly.
.PARAMETER Round
Use this value to round a sum, especially if you are formatting it to something like KB.
.PARAMETER Sort
You have the option of sorting results when using -Count or -Sum. You can sort on the value, i.e. count or sum, or on the property name. 
The default sort option is none but you can specify Ascending or Descending.
.PARAMETER SortKey
Specify if you want to sort on the value or property name. The default is Value. This parameter has no effect unless you also use -Sort.
.EXAMPLE
PS C:\> $svc="Lanmanserver","Wuauserv","DNS","ADWS"
PS C:\> $computers="chi-dc01","chi-dc02","chi-dc04"
PS C:\> $data = Get-Service -name $svc -ComputerName $computers
PS C:\> New-PSPivotTable $data -ylabel Computername -yProperty Machinename -xlabel Name -xproperty Status -verbose | format-table -autosize

Computername    ADWS     DNS Lanmanserver Wuauserv
------------    ----     --- ------------ --------
chi-dc01     Running Running      Running  Running
chi-dc02     Running Stopped      Running  Running
chi-dc04     Running Running      Running  Stopped


Create a table that shows the status of each service on each computer. The yLabel parameter renames the property so that instead of Machinename it shows Computername. 
The xLabel is the property name to analyze, in this case the service name. The xProperty value of each service becomes the table value.
.EXAMPLE
PS C:\> $files = dir c:\scripts -include *.ps1,*.txt,*.zip,*.bat -recurse
PS C:\> New-PSPivotTable $files -yProperty Directory -xLabel Extension -count | format-table -auto 

Directory                                        .ZIP .BAT .PS1 .TXT
---------                                        ---- ---- ---- ----
C:\scripts\AD-Old\New                               0    0    1    1
C:\scripts\AD-Old                                   1    0   82    1
C:\scripts\ADTFM-Scripts\LocalUsersGroups           0    0    8    0
C:\scripts\ADTFM-Scripts                            0    0   55    3
C:\scripts\en-US                                    0    0    1    0
C:\scripts\GPAE                                     0    0    8    3
C:\scripts\modhelp                                  1    0    0    0
C:\scripts\PowerShellBingo                          0    0    4    0
C:\scripts\PS-TFM                                   1    0   69    2
C:\scripts\PSVirtualBox                             0    0    0    1
C:\scripts\quark                                    0    0    0    1
C:\scripts\Toolmaking                               0    0   48    0
C:\scripts                                         55   13 1133  305

Display a table report that shows the count of each file type in each directory.

PS C:\> New-PSPivotTable $files -yProperty Directory -xLabel Extension -count | ConvertTo-HTML -title "Script Report" -CssUri C:\scripts\blue.css -PreContent "<H3>C:\Scripts</H3>" -PostContent "<H6>$(Get-Date)</H6>" | Out-File C:\work\Scripts.htm -Encoding ascii

Create a pivot table similar to the example above and create an HTML report.
.EXAMPLE
PS C:\Scripts> $files = dir -path c:\scripts\*.ps*,*.txt,*.zip,*.bat
PS C:\Scripts> New-PSPivotTable $files -yProperty Directory -xlabel Extension -Sum Length -round 2 -format kb | format-table -auto 

Directory  .PS1  .PSM1 .PS1XML .PSSC  .PSD1     .TXT    .ZIP   .BAT
---------  ----  ----- ------- -----  -----     ----    ----   ----
C:\scripts 8542 500.88  137.82 11.95  9.16  22473.86 2402.63  26.32

Analyse files by extension, measuring the total size of each extension. The value is formatted as KB to 2 decimal points.

.EXAMPLE
PS C:\scripts> New-PSPivotTable $files -yProperty Directory -xLabel Extension -Count -Sort Ascending

Directory : C:\scripts
 PSSC     : 3
 PSD1     : 7
 BAT      : 17
 PS1XML   : 24
 PSM1     : 50
 ZIP      : 74
 TXT      : 443
 PS1      : 2077

Process the collection of script files and analyze by the count of each file type. The result is sorted by the count value in ascending order.
Note that the actual output would include the period as part of the extension.
.EXAMPLE
PS C:\> $path = "\\chi-fp02\shared"

Define a variable for a path to be analyzed.

PS C:\> $files = dir $path -recurse -File | Select *, @{Name="Age";Expression={(Get-Date)-$_.LastWriteTime}},
@{Name="Bucket";Expression={
Switch([int]((Get-Date)-$_.LastWriteTime).TotalDays) {
{$_ -gt 365} {'365Plus' ; Break}
{$_ -gt 180 -AND $_ -le 365} {'1Yr' ; Break}
{$_ -gt 90 -AND $_ -le 180} {'6Mo' ; Break}
{$_ -gt 30 -AND $_ -le 90} {'3Mo' ; Break}
{$_ -gt 7 -AND $_ -le 30} { '1Mo'; Break }
{$_ -gt 0 -AND $_ -le 7} { '1Wk' ; Break }
Default { 'Today' }
} #switch
}}

Get all files and include some aging information based on the last write time.

PS C:\> New-PSPivotTable $files -yProperty Directory -xLabel Bucket -count | Out-GridView -title "File Aging"

Create a pivot table on the directory and aging buckets and display results with Out-Gridview.
.EXAMPLE
PS C:\> $data = get-eventlog system -newest 1000

Get 1000 recent events from the System eventlog.

PS C:\> New-PSPivotTable -Data $data -Count -yProperty EntryType -xLabel Source | Out-Gridview -title "System Sources"

Create a pivot table with a Y column of Entry Type and the X axis labels of the different sources. The value under each column will be the total count of entries by source. The results are piped to Out-Gridview for viewing and further sorting or filtering.

PS C:\> $e = ($data).Where({$_.EntryType -eq 'error'}) 

Create a variable with only error entries.

PS C:\> New-PSPivotTable $e -yProperty EntryType -xLabel Source -count -sort Descending


EntryType : Error
SCHANNEL  : 36
DCOM      : 23
NTFS      : 5
KERBEROS  : 1
DISK      : 1

Create a pivot table on the error source, sorted by count in descending order.

PS C:\> $k = ($data).Where({$_.source -match 'kernel'})

Create a variable of all entries where the source includes 'kernel' in the name.

PS C:\> New-PSPivotTable $k -yProperty EntryType -xLabel Source -count -sort Ascending -SortKey Name


EntryType                                : Warning
MICROSOFT-WINDOWS-KERNEL-BOOT            : 0
MICROSOFT-WINDOWS-KERNEL-GENERAL         : 0
MICROSOFT-WINDOWS-KERNEL-PNP             : 36
MICROSOFT-WINDOWS-KERNEL-POWER           : 0
MICROSOFT-WINDOWS-KERNEL-PROCESSOR-POWER : 48

EntryType                                : Information
MICROSOFT-WINDOWS-KERNEL-BOOT            : 49
MICROSOFT-WINDOWS-KERNEL-GENERAL         : 42
MICROSOFT-WINDOWS-KERNEL-PNP             : 0
MICROSOFT-WINDOWS-KERNEL-POWER           : 10
MICROSOFT-WINDOWS-KERNEL-PROCESSOR-POWER : 28

Create a pivot table for each entry type showing the count of each source. The results are sorted by the source name.
.NOTES
NAME:     New-PSPivotTable
AUTHOR:   Jeffery Hicks (@JeffHicks)
VERSION:  2.1
LASTEDIT: 30 September 2015 

This function was first published and described at http://jdhitsolutions.com/blog/powershell/2434/powershell-pivot-tables/

Learn more about PowerShell:
http://jdhitsolutions.com/blog/essential-powershell-resources/

  ****************************************************************
  * DO NOT USE IN A PRODUCTION ENVIRONMENT UNTIL YOU HAVE TESTED *
  * THOROUGHLY IN A LAB ENVIRONMENT. USE AT YOUR OWN RISK.  IF   *
  * YOU DO NOT UNDERSTAND WHAT THIS SCRIPT DOES OR HOW IT WORKS, *
  * DO NOT USE IT OUTSIDE OF A SECURE, TEST SETTING.             *
  ****************************************************************

Thanks to kdoblosky for contributing to this module.
.LINK
Measure-Object
Group-Object
Select-Object

#>

[cmdletbinding(DefaultParameterSetName = "Property")]

Param(
[Parameter(
    Position = 0,
    Mandatory,
    HelpMessage = "What is the data to analyze?"
)]
[ValidateNotNullorEmpty()]
[object]$Data,

[Parameter()]
[String]$yLabel,

[Parameter(
    Mandatory,
    HelpMessage = "What is the Y axis property?"
)]
[ValidateNotNullorEmpty()]
[String]$yProperty,

[Parameter(
    Mandatory,
    HelpMessage = "What is the X axis label?"
)]
[ValidateNotNullorEmpty()]
[string]$xLabel,

[Parameter(
    ParameterSetName = "Property"
)]
[string]$xProperty,

[Parameter(
    ParameterSetName = "Count"
)]
[switch]$Count,

[Parameter(
    ParameterSetName = "Sum"
)]
[string]$Sum,

[Parameter(
    ParameterSetName = "Sum"
)]
[ValidateSet("None","KB","MB","GB","TB","PB")]
[string]$Format = "None",

[Parameter(
    ParameterSetName = "Sum"
)]
[ValidateScript({$_ -gt 0})]
[int]$Round,

[Parameter(
    ParameterSetName = "Sum"
)]
[Parameter(
    ParameterSetName = "Count"
)]
[ValidateSet("None","Ascending","Descending")]
[string]$Sort = "None",

[Parameter(
    ParameterSetName = "Sum"
)]
[Parameter(
    ParameterSetName = "Count"
)]
[ValidateSet("Name","Value")]
[string]$SortKey = "Value"
)

Begin {
    Write-Verbose "Starting $($myinvocation.mycommand)"

    #define some values for Write-Progress
    $Activity = "PowerShell Pivot Table"
    $status = "Creating new table"
    Write-Progress -Activity $Activity -Status $Status

    #initialize an array to hold results
    $result = @()
    #if no yLabel then use yProperty name
    if (-Not $yLabel) {
        $yLabel = $yProperty
    }
    Write-Verbose "Vertical axis label is $ylabel"
} #begin

Process {    
    Write-Progress -Activity $Activity -status "Pre-Processing"
    Write-Verbose "Creating a unique list based on $xLabel"
    <#
      Filter out null values, but not blanks. Uniqueness is case sensitive 
	  so we first do a quick filtering with Select-Object, then turn each 
	  of them to upper case and finally get unique uppercase items. 
    #>
    $unique = $Data | Where {$_.$xlabel -ne $Null} | 
     Select-Object -ExpandProperty $xLabel -unique | foreach {
       $_.ToUpper()} | Select-Object -unique
         
    Write-Verbose ($unique -join  ',' | out-String).Trim()
         
    Write-Verbose "Grouping objects on $yProperty"
    Write-Progress -Activity $Activity -status "Pre-Processing" -CurrentOperation "Grouping by $yProperty"

    $grouped = $Data | Group -Property $yProperty
    $status = "Analyzing data"  
    $i = 0
    $groupcount = ($grouped | measure).count
    
    foreach ($item in $grouped ) {
      Write-Verbose "Item $($item.name)"
      $i++
      #calculate what percentage is complete for Write-Progress
      $percent = ($i/$groupcount)*100
      Write-Progress -Activity $Activity -Status $Status -CurrentOperation $($item.Name) -PercentComplete $percent
         
      $hash = [ordered]@{}
      
      #process each group
        #Calculate value depending on parameter set
        Switch ($pscmdlet.parametersetname) {
        
        "Property" {
                    Write-Verbose "Processing $xLabel for $xProperty"
                    <#
                      take each property name from the horizontal axis and make 
                      it a property name. Use the grouped property value as the 
                      new value
                    #>
                    
                    #find non-matching labels and set value to Null
                    #make each name upper case
                     $labelGroup = $item.group | Group-Object -Property $xLabel 
                     $diff = $labelGroup | Select-Object -ExpandProperty Name -unique | 
                     Foreach { $_.ToUpper() } | Select-Object -unique
                     
                     #compare the master list of unique labels with what is in this group
                     Compare-Object -ReferenceObject $Unique -DifferenceObject $diff | 
                     Select-Object -ExpandProperty InputObject | foreach {
                        #add each item and set the value to null
                        Write-Verbose "Setting $_ to null"
                       $hash.Add($_,$null)
                     }
                    
                     $item.group | foreach {
                        $v = $_.$xProperty
                        Write-Verbose "Adding $($_.$xLabel) with a value of $v"
                        $hash.Add($($_.$xLabel),$v)
                      } #foreach
                    } #property
        "Count"  {
                    Write-Verbose "Calculating count based on $xLabel"
                     $labelGroup = $item.group | Group-Object -Property $xLabel 
                     #find non-matching labels and set count to 0
                     Write-Verbose "Finding 0 count entries"
                     #make each name upper case
                     $diff = $labelGroup | Select-Object -ExpandProperty Name -unique | 
                     Foreach { $_.ToUpper() } | Select-Object -unique
                     
                     #compare the master list of unique labels with what is in this group
                     Compare-Object -ReferenceObject $Unique -DifferenceObject $diff | 
                     Select-Object -ExpandProperty InputObject | foreach {
                        #add each item and set the value to 0
                        Write-Verbose "Setting $_ to 0"
						
						# Account for blank entries
						If ([String]::IsNullOrEmpty($_)) {
							$_ = "[NONE]"
						}

                        $hash.add($_,0)
                     }
                     
                     Write-Verbose "Counting entries"
                     $labelGroup | foreach {
                        $n = ($_.name).ToUpper()
                        Write-Verbose "$n = $($_.count)"
						
						# Account for blank names
						If ([String]::IsNullOrEmpty($n)) {
							$n = "[NONE]"
						}

                        $hash.Add($n,$_.count)

                    } #foreach
                 } #count
        "Sum"  {
                    Write-Verbose "Calculating sum based on $xLabel using $sum"
                    $labelGroup = $item.group | Group-Object -Property $xLabel 
                 
                     #find non-matching labels and set count to 0
                     Write-Verbose "Finding 0 count entries"
                     #make each name upper case
                     $diff = $labelGroup | Select-Object -ExpandProperty Name -unique | 
                     Foreach { $_.ToUpper() } | Select-Object -unique
                     
                     #compare the master list of unique labels with what is in this group
                     Compare-Object -ReferenceObject $Unique -DifferenceObject $diff | 
                     Select-Object -ExpandProperty InputObject | foreach {
                        #add each item and set the value to 0
                        Write-Verbose "Setting $_ sum to 0"
						
						# Account for blank entries
						If ([String]::IsNullOrEmpty($_)) {
							$_ = "[NONE]"
						}
                        $hash.add($_,0)
                     }
                     
                     Write-Verbose "Measuring entries"
                     $labelGroup | foreach {
                        $n = ($_.name).ToUpper()
                        Write-Verbose "Measuring $n"
                        
                        $measure = $_.Group | Measure-Object -Property $Sum -sum
                        if ($Format -eq "None") {
                            $value = $measure.sum
                        }
                        else {
                            Write-Verbose "Formatting to $Format"
                             $value = $measure.sum/"1$Format"
                            }
                        if ($Round) {
                            Write-Verbose "Rounding to $Round places"
                            $Value = [math]::Round($value,$round)
                        }
                       
					   # Account for blank names
					    If ([String]::IsNullOrEmpty($n)) {
							$n = "[NONE]"
						}
                        $hash.add($n,$value)
                    } #foreach
                   
                } #Sum       
        } #switch
       
       #sort as necessary
       if ($Sort -ne "None") {
        
        Write-Verbose "Sorting order $sort"
        Write-Verbose "Sorting on $sortkey"
        #define a hashtable of parameters for Sort-Object
        $sortParams = @{Property=$SortKey}

        if ($sort -eq "Descending") {
            $sortParams.Add("Descending",$True)
        }   
        
        $sorted = $hash.GetEnumerator() | Sort-Object @sortParams
        
        #rebuild the hash table based on sorting
        $sorted | foreach -Begin {$hash = [ordered]@{}} -process { $hash.add($_.name,$_.value)}   
       }
       
       #add ylabel
       $hash.Insert(0,$yLabel,$item.name)
       
       #add each object to the results array
       Write-Verbose "Adding object to the results array"
       $result += [pscustomobject]$hash
    } #foreach item

} #process

End {

    Write-Verbose "Writing results to the pipeline"
    $result
    Write-Verbose "Ending $($myinvocation.mycommand)"
    Write-Progress -Completed -Activity $Activity -Status "Ending"

} #end

} #end function

#define an optional alias
Set-Alias -Name npt -Value New-PSPivotTable

Export-ModuleMember -Function * -Alias *


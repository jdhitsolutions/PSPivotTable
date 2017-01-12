#requires -version 4.0

Function New-PSPivotTable {

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
    Write-Verbose "Using parameter set $($pscmdlet.parameterSetName)"
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
      Also explictly convert to values to strings. 
    #>
    $unique = $Data | Where {$_.$xlabel -ne $Null} | 
     Select-Object -ExpandProperty $xLabel -unique | foreach {
       $_.ToString().ToUpper()} | Select-Object -unique
         
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


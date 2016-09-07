# PSPivotTable
A PowerShell function to create an Excel-like Pivot table in the PowerShell console. This command takes the result of a PowerShell expression and creates a pivot table object. You can use this object to analyze data patterns. For example, you could get a directory listing and then prepare a table showing the size of different file extensions for each folder.

## SYNTAX

    New-PSPivotTable [-Data] <Object> [-yLabel <String>] -yProperty <String> -xLabel <String> [-xProperty <String>] [<CommonParameters>]
    New-PSPivotTable [-Data] <Object> [-yLabel <String>] -yProperty <String> -xLabel <String> [-Count] [-Sort <String>] [-SortKey <String>] [<CommonParameters>]
    New-PSPivotTable [-Data] <Object> [-yLabel <String>] -yProperty <String> -xLabel <String> [-Sum <String>] [-Format <String>] [-Round <Int32>] [-Sort <String>] [-SortKey <String>] [<CommonParameters>]

## EXAMPLE

    PS C:\>$svc="Lanmanserver","Wuauserv","DNS","ADWS"
    PS C:\> $computers="chi-dc01","chi-dc02","chi-dc04"
    PS C:\> $data = Get-Service -name $svc -ComputerName $computers
    PS C:\> new-pspivottable $data -ylabel Computername -yProperty Machinename -xlabel Name -xproperty Status -verbose | format-table -autosize
    
    Computername    ADWS     DNS Lanmanserver Wuauserv
    ------------    ----     --- ------------ --------
    chi-dc01     Running Running      Running  Running
    chi-dc02     Running Stopped      Running  Running
    chi-dc04     Running Running      Running  Stopped
    
Create a table that shows the status of each service on each computer. The yLabel parameter renames the property so that instead of Machinename it shows Computername. The xLabel is the property name to analyze, in this case the service name. The xProperty value of each service becomes the table value.

_Last updated: September 7, 2016_
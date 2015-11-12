
Function ConvertTo-AdvHTML
{   <#
    .SYNOPSIS
        Advanced replacement of ConvertTo-HTML cmdlet
    .DESCRIPTION
        This function allows for vastly greater control over cells and rows
        in a HTML table.  It takes ConvertTo-HTML to a whole new level!  You
        can now specify what color a cell or row is (either dirctly or through 
        the use of CSS).  You can add links, pictures and pictures AS links.
        You can also specify a cell to be a bar graph where you control the 
        colors of the graph and text that can be included in the graph.
        
        All color functions are through the use of imbedded text tags inside the
        properties of the object you pass to this function.  It is important to note 
        that this function does not do any processing for you, you must make sure all 
        control tags are already present in the object before passing it to the 
        function.
        
        Here are the different tags available:
        
        Syntax                          Comment
        ===================================================================================
        [cell:<color>]<optional text>   Designate the color of the cell.  Must be 
                                        at the beginning of the string.
                                        Example:
                                            [cell:red]System Down
                                            
        [row:<color>]                   Designate the color of the row.  This control
                                        can be anywhere, in any property of the object.
                                        Example:
                                            [row:orchid]
                                            
        [cellclass:<class>]<optional text>  
                                        Designate the color, and other properties, of the
                                        cell based on a class in your CSS.  You must 
                                        have the class in your CSS (use the -CSS parameter).
                                        Must be at the beginning of the string.
                                        Example:
                                            [cellclass:highlight]10mb
                                            
        [rowclass:<class>]              Designate the color, and other properties, of the
                                        row based on a class in your CSS.  You must 
                                        have the class in your CSS (use the -CSS parameter).
                                        This control can be anywhere, in any property of the 
                                        object.
                                        Example:
                                            [rowclass:greyishbold]
                                            
        [image:<height;width;url>]<alternate text>
                                        Include an image in your cell.  Put size of picture
                                        in pixels and url seperated by semi-colons.  Format
                                        must be height;width;url.  You can also include other
                                        text in the cell, but the [image] tag must be at the
                                        end of the tag (so the alternate text is last).
                                        Example:
                                            [image:100;200;http://www.sampleurl.com/sampleimage.jpg]Alt Text For Image
                                            
        [link:<url>]<link text>         Include a link in your cell.  Other text is allowed in
                                        the string, but the [link] tag must be at the end of the 
                                        string.
                                        Example:
                                            blah blah blah [link:www.thesurlyadmin.com]Cool PowerShell Link
                                            
        [linkpic:<height;width;url to pic>]<url for link>
                                        This tag uses a picture which you can click on and go to the
                                        specified link.  You must specify the size of the picture and 
                                        url where it is located, this information is seperated by semi-
                                        colons.  Other text is allowed in the string, but the [link] tag 
                                        must be at the end of the string.
                                        Example:
                                            [linkpic:100;200;http://www.sampleurl.com/sampleimage.jpg]www.thesurlyadmin.com
                                            
        [bar:<percent;bar color;remainder color>]<optional text>
                                        Bar graph makes a simple colored bar graph within the cell.  The
                                        length of the bar is controlled using <percent>.  You can 
                                        designate the color of the bar, and the color of the remainder
                                        section.  Due to the mysteries of HTML, you must designate a 
                                        width for the column with the [bar] tag using the HeadWidth parameter.
                                        
                                        So if you had a percentage of 95, say 95% used disk you
                                        would want to highlight the remainder for your report:
                                        Example:
                                            [bar:95;dark green;red]5% free
                                        
                                        What if you were at 30% of a sales goal with only 2 weeks left in
                                        the quarter, you would want to highlight that you have a problem.
                                        Example:
                                            [bar:30;darkred;red]30% of goal
    .PARAMETER InputObject
        The object you want converted to an HTML table
    .PARAMETER HeadWidth
        You can specify the width of a cell.  Cell widths are in pixels
        and are passed to the parameter in array format.  Each element
        in the array corresponds to the column in your table, any element
        that is set to 0 will designate the column with be dynamic.  If you had
        four elements in your InputObject and wanted to make the 4th a fixed
        width--this is required for using the [bar] tag--of 600 pixels:
        
        -HeadWidth 0,0,0,600
    .PARAMETER CSS
        Designate custom CSS for your HTML
    .PARAMETER Title
        Specifies a title for the HTML file, that is, the text that appears between the <TITLE> tags.
    .PARAMETER PreContent
        Specifies text to add before the opening <TABLE> tag. By default, there is no text in that position.
    .PARAMETER PostContent
        Specifies text to add after the closing </TABLE> tag. By default, there is no text in that position.
    .PARAMETER Body
        Specifies the text to add after the opening <BODY> tag. By default, there is no text in that position.
    .PARAMETER Fragment
        Generates only an HTML table. The HTML, HEAD, TITLE, and BODY tags are omitted.
    .INPUTS
        System.Management.Automation.PSObject
        You can pipe any .NET object to ConvertTo-AdvHtml.
    .OUTPUTS
        System.String
        ConvertTo-AdvHtml returns series of strings that comprise valid HTML.
    .EXAMPLE
        $Data = @"
Server,Description,Status,Disk
[row:orchid]Server1,Hello1,[cellclass:up]Up,"[bar:45;Purple;Orchid]55% Free"
Server2,Hello2,[cell:green]Up,"[bar:65;DarkGreen;Green]65% Used"
Server3,Goodbye3,[cell:red]Down,"[bar:95;DarkGreen;DarkRed]5% Free"
server4,This is quite a cool test,[cell:green]Up,"[image:150;650;http://pughspace.files.wordpress.com/2014/01/test-connection.png]Test Images"
server5,SurlyAdmin,[cell:red]Down,"[link:http://thesurlyadmin.com]The Surly Admin"
server6,MoreSurlyAdmin,[cell:purple]Updating,"[linkpic:150;650;http://pughspace.files.wordpress.com/2014/01/test-connection.png]http://thesurlyadmin.com"
"@
        $Data = $Data | ConvertFrom-Csv
        $HTML = $Data | ConvertTo-AdvHTML -HeadWidth 0,0,0,600 -PreContent "<p><h1>This might be the best report EVER</h1></p><br>" -PostContent "<br>Done! $(Get-Date)" -Title "Cool Test!"
        
        This is some sample code where I try to put every possibile tag and use into a single set
        of data.  $Data is the PSObject 4 columns.  Default CSS is used, so the [cellclass:up] tag
        will not work but I left it there so you can see how to use it.
    .NOTES
        Author:             Martin Pugh
        Twitter:            @thesurlyadm1n
        Spiceworks:         Martin9700
        Blog:               www.thesurlyadmin.com
          
        Changelog:
            1.0             Initial Release
    .LINK
        http://thesurlyadmin.com/convertto-advhtml-help/
    .LINK
        http://community.spiceworks.com/scripts/show/2448-create-advanced-html-tables-in-powershell-convertto-advhtml
    #>
    #requires -Version 2.0
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true,
            ValueFromPipeline=$true)]
        [Object[]]$InputObject,
        [string[]]$HeadWidth,
        [string]$CSS = @"
<style>
TABLE {border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}
TH {border-width: 1px;padding: 3px;border-style: solid;border-color: black;background-color: #6495ED;font-size:120%;}
TD {border-width: 1px;padding: 3px;border-style: solid;border-color: black;}
</style>
"@,
        [string]$Title,
        [string]$PreContent,
        [string]$PostContent,
        [string]$Body,
        [switch]$Fragment
    )
    
    Begin {
        If ($Title)
        {   $CSS += "`n<title>$Title</title>`n"
        }
        $Params = @{
            Head = $CSS
        }
        If ($PreContent)
        {   $Params.Add("PreContent",$PreContent)
        }
        If ($PostContent)
        {   $Params.Add("PostContent",$PostContent)
        }
        If ($Body)
        {   $Params.Add("Body",$Body)
        }
        If ($Fragment)
        {   $Params.Add("Fragment",$true)
        }
        $Data = @()
    }
    
    Process {
        ForEach ($Line in $InputObject)
        {   $Data += $Line
        }
    }
    
    End {
        $Html = $Data | ConvertTo-Html @Params

        $NewHTML = @()
        ForEach ($Line in $Html)
        {   If ($Line -like "*<th>*")
            {   If ($Headwidth)
                {   $Index = 0
                    $Reg = $Line | Select-String -AllMatches -Pattern "<th>(.*?)<\/th>"
                    ForEach ($th in $Reg.Matches)
                    {   If ($Index -le ($HeadWidth.Count - 1))
                        {   If ($HeadWidth[$Index] -and $HeadWidth[$Index] -gt 0)
                            {   $Line = $Line.Replace($th.Value,"<th style=""width:$($HeadWidth[$Index])px"">$($th.Groups[1])</th>")
                            }
                        }
                        $Index ++
                    }
                }
            }
        
            Do {
                Switch -regex ($Line)
                {   "<td>\[cell:(.*?)\].*?<\/td>"
                    {   $Line = $Line.Replace("<td>[cell:$($Matches[1])]","<td style=""background-color:$($Matches[1])"">")
                        Break
                    }
                    "\[cellclass:(.*?)\]"
                    {   $Line = $Line.Replace("<td>[cellclass:$($Matches[1])]","<td class=""$($Matches[1])"">")
                        Break
                    }
                    "\[row:(.*?)\]"
                    {   $Line = $Line.Replace("<tr>","<tr style=""background-color:$($Matches[1])"">")
                        $Line = $Line.Replace("[row:$($Matches[1])]","")
                        Break
                    }
                    "\[rowclass:(.*?)\]"
                    {   $Line = $Line.Replace("<tr>","<tr class=""$($Matches[1])"">")
                        $Line = $Line.Replace("[rowclass:$($Matches[1])]","")
                        Break
                    }
                    "<td>\[bar:(.*?)\](.*?)<\/td>"
                    {   $Bar = $Matches[1].Split(";")
                        $Width = 100 - [int]$Bar[0]
                        If (-not $Matches[2])
                        {   $Text = "&nbsp;"
                        }
                        Else
                        {   $Text = $Matches[2]
                        }
                        $Line = $Line.Replace($Matches[0],"<td><div style=""background-color:$($Bar[1]);float:left;width:$($Bar[0])%"">$Text</div><div style=""background-color:$($Bar[2]);float:left;width:$width%"">&nbsp;</div></td>")
                        Break
                    }
                    "\[image:(.*?)\](.*?)<\/td>"
                    {   $Image = $Matches[1].Split(";")
                        $Line = $Line.Replace($Matches[0],"<img src=""$($Image[2])"" alt=""$($Matches[2])"" height=""$($Image[0])"" width=""$($Image[1])""></td>")
                    }
                    "\[link:(.*?)\](.*?)<\/td>"
                    {   $Line = $Line.Replace($Matches[0],"<a href=""$($Matches[1])"">$($Matches[2])</a></td>")
                    }
                    "\[linkpic:(.*?)\](.*?)<\/td>"
                    {   $Images = $Matches[1].Split(";")
                        $Line = $Line.Replace($Matches[0],"<a href=""$($Matches[2])""><img src=""$($Image[2])"" height=""$($Image[0])"" width=""$($Image[1])""></a></td>")
                    }
                    Default
                    {   Break
                    }
                }
            } Until ($Line -notmatch "\[.*?\]")
            $NewHTML += $Line
        }
        Return $NewHTML
    }
}

<#
cls
$Data = @"
Server,Description,Status,Disk
[row:orchid]Server1,Hello1,[cellclass:up]Up,"[bar:45;Purple;Orchid]55% Free"
Server2,Hello2,[cell:green]Up,"[bar:65;DarkGreen;Green]Hello Laura"
Server3,Goodbye3,[cell:red]Down,"[bar:95;DarkGreen;DarkRed]5% Free"
server4,This is quite a cool test,[cell:green]Up,"[image:150;650;http://pughspace.files.wordpress.com/2014/01/test-connection.png]Test Images"
server5,SurlyAdmin,[cell:red]Down,"[link:http://thesurlyadmin.com]The Surly Admin"
server6,MoreSurlyAdmin,[cell:purple]Updating,"[linkpic:150;650;http://pughspace.files.wordpress.com/2014/01/test-connection.png]http://thesurlyadmin.com"
"@
$Data = $Data | ConvertFrom-Csv#>
$Data = @(
    [PSCustomObject]@{
        Server = "[row:orchid]Server1"
        Description = "Hello1"
        Status = "[cellclass:up]Up"
        Disk = "[bar:45;Purple;Orchid]55% Free"
    },
    [PSCustomObject]@{
        Server = "Server2"
        Description = "Hello2"
        Status = "[cell:green]Up"
        Disk = "[bar:65;DarkGreen;Green]65% Used"
    },
    [PSCustomObject]@{
        Server = "Server3"
        Description = "Goodbye3"
        Status = "[cell:red]Down"
        Disk = "[bar:95;DarkGreen;DarkRed]5% Free"
    },
    [PSCustomObject]@{
        Server = "Server4"
        Description = "This is quite a cool test"
        Status = "[cell:green]Up"
        Disk = "[image:150;650;http://pughspace.files.wordpress.com/2014/01/test-connection.png]Test Images"
    },
    [PSCustomObject]@{
        Server = "Server5"
        Description = "SurlyAdmin"
        Status = "[cell:red]Down"
        Disk = "[link:http://thesurlyadmin.com]The Surly Admin"
    },
    [PSCustomObject]@{
        Server = "Server6"
        Description = "MoreSurlyAdmin"
        Status = "[cell:purple]Updating"
        Disk = "[linkpic:150;650;http://pughspace.files.wordpress.com/2014/01/test-connection.png]http://thesurlyadmin.com"
    }
)

$HTML = $Data | ConvertTo-AdvHTML -HeadWidth 0,0,0,600 -PreContent "<p><h1>ConvertTo-AdvHTML Sample Report</h1></p><br>" -PostContent "<br>Done! $(Get-Date)" -Title "Cool Test!"


$Html | Out-File c:\Dropbox\Test\test.html
& c:\Dropbox\Test\test.html
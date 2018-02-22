
 $folderpath    = 'C:\temp\ReceivedEmails\*'
 $pageCountFile = 'C:\temp\PageCount.csv'
 $fileTypes     = '*.pptx','*.ppt'

Function Set-Variables
{
 $confirmConversion = $false
 $readOnly = [Microsoft.Office.Core.MsoTriState]::msoFalse
 $untitled = [Microsoft.Office.Core.MsoTriState]::msoFalse
 $withWindow = [Microsoft.Office.Core.MsoTriState]::msoFalse
 $numberOfPages = 0
 Set-OutputFile
} #end Set-Variables

Function Set-OutputFile
{
 if(Test-Path -path $pageCountFile)
   { Remove-Item -path $pageCountFile }
 'name,pageCount' >> $pageCountFile
 Get-WordDocuments
} #end Set-OutputFile

Function Get-WordDocuments
{
  'Counting pages in Word Docs in ' + $folderPath
 $powerpoint = New-Object -ComObject powerpoint.application
 Get-ChildItem -path $folderpath -include $fileTypes -Recurse |
 foreach-object `
  {
   $path =  ($_.fullname).substring(0,($_.FullName).lastindexOf('.'))
   $presentation = $powerpoint.presentations.open($_.fullname, $readOnly, $untitled, $withWindow)
   $slidesCount = $presentation.Slides.Count
   $($_.name) + ',' + $($slidesCount)  >> $pageCountFile
   $presentation.close()
  } #end Foreach-Object
 $powerpoint.Quit()
 Get-pageCount
} #end Get-WordDocuments

Function Get-pageCount
{
 $wdcsv = import-csv -path $pageCountFile
 if (-Not $wdcsv.length) { $numberOfPages = [int32]$wdcsv.pageCount }
    else {
        for ($i = 0 ; $i -le $wdcsv.length -1 ; $i++) {
            $numberOfPages += [int32]$wdcsv[$i].pageCount
        }
    }
 $numberOfPages
} #end Get-pageCount

# *** Entry Point to Script ***

Set-Variables

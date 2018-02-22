# https://blogs.technet.microsoft.com/heyscriptingguy/2009/10/13/hey-scripting-guy-how-do-i-count-the-number-of-pages-in-a-group-of-office-word-documents/

Function Set-Variables
{
 $folderpath = "c:fso*"
 $fileTypes = "*.docx","*doc"
 $confirmConversion = $false
 $readOnly = $true
 $addToRecent = $false
 $passwordDocument = "password"
 $pageCountFile = "C:fsoPageCount.csv"
 $numberOfPages = 0
 Set-OutputFile
} #end Set-Variables

Function Set-OutputFile
{
 if(Test-Path -path $pageCountFile)
   { Remove-Item -path $pageCountFile }
 "name,pageCount" >> $pageCountFile
 Get-WordDocuments
} #end Set-OutputFile

Function Get-WordDocuments
{
  "Counting pages in Word Docs in $folderPath"
 $word = New-Object -ComObject word.application
 $word.visible = $false
 Get-ChildItem -path $folderpath -include $fileTypes |
 foreach-object `
  {
   $path =  ($_.fullname).substring(0,($_.FullName).lastindexOf("."))
   $doc = $word.documents.open($_.fullname, $confirmConversion, $readOnly, 
   $addToRecent,   $passwordDocument)
   $window = $doc.ActiveWindow
   $panes = $window.Panes
   $pane = $Panes.item(1)
   "  $($_.name), $($pane.pages.count)"  >> $pageCountFile
   $doc.close()
  } #end Foreach-Object
 $word.Quit()
 Get-pageCount
} #end Get-WordDocuments

Function Get-pageCount
{
 $wdcsv = import-csv -path $pageCountFile
 for ($i = 0 ; $i -le $wdcsv.length -1 ; $i++)
 {
  $numberOfPages += [int32]$wdcsv[$i].pageCount
 }
 $numberOfPages
} #end Get-pageCount

# *** Entry Point to Script ***

Set-Variables

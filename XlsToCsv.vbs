'* script converts xls files to csv files'
'* run on command line: XlsToCsv.vbs [sourcexlsFile].xls [destinationcsvfile].csv'
'* http://stackoverflow.com/questions/1858195/convert-xls-to-csv-on-command-line'

csv_format = 6

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objArgs = WScript.Arguments
strFile= objArgs(0)

src_file = objFSO.GetAbsolutePathName(strFile)

'* get parent directory to create temp csv file'
'* https://www.experts-exchange.com/questions/24286795/VBScript-get-self-and-parent-folder-name.html'

dim CurrentDirectory
CurrentDirectory = objFso.GetParentFolderName(strFile)
dest_file = CurrentDirectory & "\tempPositivePay.csv"

Dim oExcel
Set oExcel = CreateObject("Excel.Application")

Dim oBook
Set oBook = oExcel.Workbooks.Open(src_file)

oBook.SaveAs dest_file, csv_format

oBook.Close False
oExcel.Quit


'* this part reads the CSV file to grab the contents'
'* http://www.tek-tips.com/viewthread.cfm?qid=1231007'
        
dim fs,objTextFile
set fs=CreateObject("Scripting.FileSystemObject")
dim arrStr
set objTextFile = fs.OpenTextFile(dest_file)

Do while NOT objTextFile.AtEndOfStream
  arrStr = split(objTextFile.ReadLine,",")
Loop

objTextFile.Close
set objTextFile = Nothing
set fs = Nothing


'* grab last item of array'
'* http://stackoverflow.com/questions/1349950/get-last-element-of-string-array-in-vb6'

dim lastValue
lastValue = arrStr(UBound(arrStr))


'* convert last value of csv (date) to mm/dd/yyyy format'
'* https://www.w3schools.com/asp/func_formatdatetime.asp'

newDate = FormatDateTime(lastValue)


'* add 0 to beginning of day and month if less that two digits'
'* http://stackoverflow.com/questions/28765980/vb-script-date-formats-yyyymmddhhmmss'

Function add0 (testIn)
 Select Case Len(testIn) < 2
   CASE TRUE
     add0 = "0" & testIn
   Case Else
     add0 = testIn
  End Select      
End Function

newMonth = add0(Month(newDate))
newDay = add0(Day(newDate))


'* get the year and last two digits of the year'
'* http://stackoverflow.com/questions/2223287/does-vbscript-have-a-substring-function'

newYear = Year(newDate)

dim shortYear
shortYear = Mid(newYear,3,2)


'* get current user'
'* http://stackoverflow.com/questions/22276361/how-to-get-username-with-vbs'

strUser = CreateObject("WScript.Network").UserName


'* put day, month, and year together to create new destination file'

newDestFile = "C:\Users\" & strUser & "\Box\Accounting\Positive Pay CSVs\" & newYear & "\" & newMonth & "." & newDay & "." & shortYear & ".csv"


'* rename file and move to Box folder'
'* http://stackoverflow.com/questions/17660117/rename-a-file-using-vbscript'

Dim Fso
Set Fso = WScript.CreateObject("Scripting.FileSystemObject")
Fso.MoveFile dest_file, newDestFile

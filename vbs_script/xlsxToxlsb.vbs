Set args = WScript.Arguments
set fs = CreateObject("Scripting.FileSystemObject")

set excel = CreateObject("Excel.Application")


for each arg in args
if LCase(fs.GetExtensionName(arg)) = "xlsx"  then       			'如果是xlsx文件
	set wb= excel.Workbooks.Open(fs.GetAbsolutePathName(arg)) 	 '打开excel文件

	SavePath=fs.GetParentFolderName(arg) &"\"& fs.GetBaseName(arg) &".xlsb"  		'拼接新的文件名

	wb.SaveAs SavePath,50,,,,False 					'另存为xlsb  51=xlsx 50=xlsb			
	
	fs.DeleteFile fs.GetAbsolutePathName(arg),true			'删除原来的xlsx文件
	
	wb.close	
end if


next
excel.Quit
set excel=nothing
set fs=nothing
set args=nothing



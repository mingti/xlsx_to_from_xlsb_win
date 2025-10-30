Set args = WScript.Arguments
set fs = CreateObject("Scripting.FileSystemObject")

set excel = CreateObject("Excel.Application")


for each arg in args
if LCase(fs.GetExtensionName(arg)) = "xlsx"  then       			'�����xlsx�ļ�
	set wb= excel.Workbooks.Open(fs.GetAbsolutePathName(arg)) 	 '��excel�ļ�

	SavePath=fs.GetParentFolderName(arg) &"\"& fs.GetBaseName(arg) &".xlsb"  		'ƴ���µ��ļ���

	wb.SaveAs SavePath,50,,,,False 					'���Ϊxlsb  51=xlsx 50=xlsb			
	
	fs.DeleteFile fs.GetAbsolutePathName(arg),true			'ɾ��ԭ����xlsx�ļ�
	
	wb.close	
end if


next
excel.Quit
set excel=nothing
set fs=nothing
set args=nothing



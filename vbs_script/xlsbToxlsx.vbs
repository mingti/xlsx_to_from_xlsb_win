Set args = WScript.Arguments
set fs = CreateObject("Scripting.FileSystemObject")

set excel = CreateObject("Excel.Application")


for each arg in args
if LCase(fs.GetExtensionName(arg)) = "xlsb"  then       		'�����xlsb�ļ�
	set wb= excel.Workbooks.Open(fs.GetAbsolutePathName(arg)) 	 '��excel�ļ�

	SavePath=fs.GetParentFolderName(arg) &"\"& fs.GetBaseName(arg) &".xlsx"  		'ƴ���µ��ļ���

	wb.SaveAs SavePath,51,,,,False 					'���Ϊxlsx  51=xlsx 50=xlsb			
	
	fs.DeleteFile fs.GetAbsolutePathName(arg),true			'ɾ��ԭ����xlsb�ļ�
	
	wb.close	
end if

next


excel.Quit
set excel=nothing
set fs=nothing
set args=nothing



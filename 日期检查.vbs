Option Explicit
Function OpenExcelApp()
	Dim obj : Set obj = Nothing
	On Error Resume Next
	Set obj = CreateObject("Excel.Application")
	If Err.Number<>0 Then
		Err.clear()
		Set obj = CreateObject("KET.Application")
	End If
	Set OpenExcelApp= obj
	Exit Function
End Function

Dim path : path = CreateObject("Scripting.FileSystemObject").GetFolder(".").Path & "\"
Dim file : file = "物流公司资质备案.xlsx" ' 在这里修改默认文件名
Dim FileDialog
Function OpenFile(app)
	Set OpenFile = Nothing
	On Error Resume Next
	Set OpenFile = objExcel.Workbooks.Open(path + file)
	If Err.Number<>0 Then
		Err.clear()
		' file = app.GetOpenFilename("Excel文件,*.xl*;*.xlsx;*.xlsm;*.xlsb;*.xlam;*.xltx;*.xltm;*.xls;*.xla;*.xlt;*.xlm;*.xlw, 所有文件, *.*", 1, "未找到"& file &"！请手动选择：", "", false)
		Set FileDialog = app.FileDialog(3) ' msoFileDialogFilePicker
		FileDialog.Filters.Add "Excel文件", "*.xl*;*.xlsx;*.xlsm;*.xlsb;*.xlam;*.xltx;*.xltm;*.xls;*.xla;*.xlt;*.xlm;*.xlw", 1
		FileDialog.Filters.Add "所有文件", "*.*", 2
		FileDialog.Title = "未找到"& file &"！请手动选择目标文件："
		FileDialog.InitialFileName = path
		FileDialog.show()
		file = FileDialog.SelectedItems(1)
		if file=false then
			app.Quit
			WScript.Quit
		end if
		' msgbox file
		Set OpenFile = objExcel.Workbooks.Open(file)
	End If
	If Err.Number<>0 Then
		Err.clear()
		msgbox "文件打开失败。"
		app.Quit
		WScript.Quit
	End If
	Exit Function
End Function

dim curDate : curDate = Date
Function CheckDate(ndt)
	CheckDate = true
	dim dif
	dif = DateDiff("d", curDate, ndt)
	if dif<0 or dif>61 then
		CheckDate = false
	end if
	Exit Function
End Function


Dim objExcel
Set objExcel = OpenExcelApp()
if objExcel Is Nothing then
	msgbox "未能找到 Microsoft Excel 或 WPS 等有效的Excel软件，请确认安装相应软件后再试"
	WScript.Quit
end if
' objExcel.Visible = true
Dim objWorkbook : Set objWorkbook = OpenFile(objExcel)


' 获取工作表
Dim Sheets : Set Sheets = objWorkbook.Worksheets
dim sheet, name, countS, countR, countC : countS = 0 : countR = 0 : countC = 0
dim list, row, col, cell, val
dim warned : warned = false
for each sheet in Sheets
	if IsEmpty(name)=False then
		' list = list & "，"
	end if
	name = sheet.Name
	list = list & "“" & name & "” "
	countS = countS+1
	for each row in sheet.UsedRange.Rows
		warned = false : row.Interior.ColorIndex = 0
		for each cell in row.Cells
			' msgbox cell.value ' merged cell?
			val = cell.MergeArea(1,1).Value
			if IsDate(val) then
				cell.Font.ColorIndex = 0 : cell.Font.Bold = false
				if CheckDate(CDate(val)) then
					if warned=false then
						warned = true : row.Interior.ColorIndex = 19
						countR = countR+1
					end if
					countC = countC+1
					cell.Interior.ColorIndex = 44 : cell.Font.ColorIndex = 3 : cell.Font.Bold = True  
				end if
			end if
		next
	next
next

if msgbox("共检查了 "& list & countS &" 个工作表。"& VbCrLf &"扫描到 "& countC &" 格临期日期，共属于 "& countR &" 行记录。" &VbCrLf &"是否查看详情？（选择“否”将保存表格中的临期高亮）", VbYesNo, "扫描完毕") = VbYes then
	Sheets(1).activate
	' objExcel.DisplayFullScreen = true ' 会隐藏工具栏
	objExcel.WindowState = -4137 ' xlMaximized
	objExcel.visible = true
else
	objWorkbook.Save
	objWorkbook.Close
	objExcel.Quit
end if
Set objExcel = Nothing : Set objWorkbook = Nothing : Set Sheets = Nothing
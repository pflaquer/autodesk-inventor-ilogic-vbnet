Sub Main()
Dim files() As String
''Conditional Prompt to Continue W/ Program
If MsgBox("Would You Like to Continue?",MessageBoxButtons.OKCancel,"DWG Exporter") = MsgBoxResult.Ok Then
''Set files to current directory
files = System.IO.Directory.GetFiles(ThisDoc.Path, "*.idw")
''Iterate Through Each File Within Directory and Push it to the Array
For Each x As String In files
	''Show Dialog
    MessageBox.Show(x, "Current DWG To Export")
	''Launch The Current Document in Iteration Referenced By X
	ThisDoc.Launch(x)
	'' Save The Document as a DWG
	ThisApplication.ActiveDocument.SaveAs(x & ".dwg", True)
	'' Close The Document Before Moving Onto The Next Document in Array
ThisApplication.ActiveDocument.Close()
Next
MsgBox("Operation Completed Succesfully!", MessageBoxIcon.Information, "Results")
Else
	MessageBox.Show("Operation Halted!", "Press OK To Exit")
End If
End Sub

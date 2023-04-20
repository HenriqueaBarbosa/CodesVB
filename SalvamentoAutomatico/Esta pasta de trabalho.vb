Private Sub Workbook_Open()
	Call StartTimer
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
	Call StopTimer
End Sub


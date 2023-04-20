Public RunWhen As Double
Public Const cRunIntervalSeconds = 60 '1 minuto
Public Const cRunWhat = "SalvamentoProgramado"
Sub StartTimer()
	RunWhen = Now + TimeSerial(0, 0, cRunIntervalSeconds)
	Application.OnTime EarliestTime:=RunWhen, Procedure:=cRunWhat, _
	Schedule:=True
End Sub
Sub SalvamentoProgramado()
	If Application.ThisWorkbook.Saved = False Then
		Application.ThisWorkbook.Save
	End If
	StartTimer
End Sub
Sub StopTimer()
	Application.OnTime EarliestTime:=RunWhen, Procedure:=cRunWhat, _
	Schedule:=False
End Sub
Public RunWhen As Double
Public Const cRunIntervalSeconds = 60 '1 minuto
Public Const cRunWhat = "SalvamentoProgramado"
Sub StartTimer()
    RunWhen = Now + TimeSerial(0, 0, cRunIntervalSeconds)
    Application.OnTime EarliestTime:=RunWhen, Procedure:=cRunWhat, _
    Schedule:=True
End Sub
Sub SalvamentoProgramado()
    Application.ThisWorkbook.Save
    StartTimer
End Sub
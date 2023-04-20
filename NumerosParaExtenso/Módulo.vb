Public Function EXTENSO(NumeroParaConverter As String) As String
	Dim sExtensoFinal As String, sExtensoAtual As String
	Dim i As Integer
	Dim iQtdGrupos As Integer
	Dim sDecimais As String
	Dim sMoedaSing As String, sMoedaPlu As String, sCentavos As String
	Dim bSufMoeda As Boolean
	
	If InStr(1, NumeroParaConverter, ",") > 0 Then
		sDecimais = Right(NumeroParaConverter, Len(NumeroParaConverter) - InStr(1, NumeroParaConverter, ","))
		NumeroParaConverter = Mid(NumeroParaConverter, 1, InStr(1, NumeroParaConverter, ",") - 1)
	End If
	
	iQtdGrupos = Fix(Len(NumeroParaConverter) / 3)
	If Len(NumeroParaConverter) Mod 3 > 0 Then
		iQtdGrupos = iQtdGrupos + 1
	End If
	
	If iQtdGrupos > 2 Then bSufMoeda = True
	For i = iQtdGrupos To 1 Step -1
		sExtensoAtual = DesmembraValor(NumeroParaConverter, i)
		If i = 1 Then
			If sExtensoAtual = "" Then
				sExtensoFinal = sExtensoFinal & sExtensoAtual
			Else
				If sExtensoFinal = "" Then
					sExtensoFinal = sExtensoFinal & sExtensoAtual
				Else
					sExtensoFinal = sExtensoFinal & " e " & sExtensoAtual
				End If
			End If
		Else
			sExtensoFinal = sExtensoFinal & sExtensoAtual
		End If
		If iQtdGrupos > 2 Then
			Select Case i
			Case 1, 2
				If sExtensoAtual <> "" Then
					bSufMoeda = False
				End If
			End Select
		End If
		Next i
		
		sMoedaPlu = " reais"
		sMoedaSing = " real"
		If bSufMoeda = True Then sMoedaPlu = " de reais"
		
		sCentavos = EscreveCentavos(sDecimais)
		
		sExtensoFinal = IIf((sExtensoFinal = ""), "", sExtensoFinal & IIf((sExtensoFinal = "um"), sMoedaSing, sMoedaPlu)) _
		& IIf((sExtensoFinal = ""), sCentavos, IIf((sCentavos = ""), "", " e " & sCentavos))
		
		EXTENSO = sExtensoFinal
	End Function
	Private Function DesmembraValor(sValor As String, iGrupoDiv As Integer) As String
		Dim iValor As Integer
		Dim sExtenso As String
		Dim iDivResto As Integer
		Dim iDivInteiro As Integer
		Dim iPosInicMid As Integer
		Dim iTamMid As Integer
		Dim sComplemento As String
		Dim vArrDez1 As Variant
		Dim vArrDez2 As Variant
		Dim vArrCentena As Variant
		vArrDez1 = Array("", "um", "dois", "três", "quatro", "cinco", "seis", "sete", "oito", "nove", _
		"dez", "onze", "doze", "treze", "quatorze", "quinze", "dezesseis", "dezessete", _
		"dezoito", "dezenove")
		vArrDez2 = Array("vinte", "trinta", "quarenta", "cinquenta", "sessenta", _
		"setenta", "oitenta", "noventa")
		vArrCentena = Array("cem", "cento", "duzentos", "trezentos", "quatrocentos", _
		"quinhentos", "seiscentos", "setecentos", "oitocentos", "novecentos")
		
		iPosInicMid = Len(sValor) - ((3 * iGrupoDiv) - 1)
		If iPosInicMid <= 1 Then
			iTamMid = 2 + iPosInicMid
		Else
			iTamMid = 3
		End If
		If iPosInicMid < 1 Then iPosInicMid = 1
		iValor = CInt(Mid(sValor, iPosInicMid, iTamMid))
		Select Case iGrupoDiv
		Case 2
			sComplemento = " mil "
		Case 3
			If iValor = 1 Then
				sComplemento = " milhão "
			Else
				sComplemento = " milhões "
			End If
		Case 4
			If iValor = 1 Then
				sComplemento = " bilhão "
			Else
				sComplemento = " bilhões "
			End If
		Case 5
			If iValor = 1 Then
				sComplemento = " trilhão "
			Else
				sComplemento = " trilhões "
			End If
		End Select
		Select Case iValor
		Case 0 To 19
			sExtenso = vArrDez1(iValor)
		Case 20 To 99
			iDivInteiro = Fix(iValor / 10)
			iDivResto = iValor Mod 10
			If iDivResto = 0 Then
				sExtenso = vArrDez2(iDivInteiro - 2)
			Else
				sExtenso = vArrDez2(iDivInteiro - 2) & " e " & vArrDez1(iDivResto)
			End If
		Case 100 To 999
			iDivInteiro = Fix(iValor / 100)
			iDivResto = iValor Mod 100
			If iDivResto = 0 Then
				If iDivInteiro = 1 Then
					sExtenso = vArrCentena(0)   
				Else
					sExtenso = vArrCentena(iDivInteiro) 
				End If
			Else
				sExtenso = vArrCentena(iDivInteiro) & " e "
				Select Case iDivResto
				Case 0 To 19
					sExtenso = sExtenso & vArrDez1(iDivResto)
				Case 20 To 99
					iDivInteiro2 = Fix(iDivResto / 10)
					iDivResto2 = iDivResto Mod 10
					If iDivResto2 = 0 Then
						sExtenso = sExtenso & vArrDez2(iDivInteiro2 - 2)
					Else
						sExtenso = sExtenso & vArrDez2(iDivInteiro2 - 2) & " e " & vArrDez1(iDivResto2)
					End If
				End Select
			End If
		End Select
		DesmembraValor = sExtenso & IIf(iValor > 0, sComplemento, "")
	End Function
	Private Function EscreveCentavos(sCent As String) As String
		Dim sExtenso As String
		Dim iDivResto As Integer
		Dim iDivInteiro As Integer
		Dim sComplemento As String
		Dim vArrDez1 As Variant
		Dim vArrDez2 As Variant
		Dim iCent As Integer
		vArrDez1 = Array("", "um", "dois", "três", "quatro", "cinco", "seis", "sete", "oito", "nove", _
		"dez", "onze", "doze", "treze", "quatorze", "quinze", "dezesseis", "dezessete", _
		"dezoito", "dezenove")
		vArrDez2 = Array("vinte", "trinta", "quarenta", "cinquenta", "sessenta", _
		"setenta", "oitenta", "noventa")
		
		iCent = Fix(sCent & String(2 - Len(sCent), "0"))
		
		If iCent = 1 Then
			sComplemento = " centavo"
		Else
			sComplemento = " centavos"
		End If
		
		Select Case iCent
		Case 0 To 19
			sExtenso = vArrDez1(iCent)
		Case 20 To 99
			iDivInteiro = Fix(iCent / 10)
			iDivResto = iCent Mod 10
			If iDivResto = 0 Then
				sExtenso = vArrDez2(iDivInteiro - 2)
			Else
				sExtenso = vArrDez2(iDivInteiro - 2) & " e " & vArrDez1(iDivResto)
			End If
		End Select
		EscreveCentavos = IIf(iCent > 0, sExtenso & sComplemento, "")
	End Function
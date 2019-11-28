Dim balance , itf , regimen

'balance 	= DataTable("balance",1)
'itf			= DataTable("itf",1)
'regimen		= DataTable("regimen",1)

balance = "Si"
itf		= "Si"
regimen = "Mype"

Call Fun_Balance(balance)
Call Fun_ITF(itf)
Call Fun_regimen(regimen)

Function Fun_Balance(balance)
		
	If balance = "Si" Then
		JavaWindow("Sistema Integrador v1").JavaTab("Paso 1:").JavaTab("Sección Informativa").JavaTab("Identificación").JavaRadioButton("Balancei_Rdb_Si").Set "ON"	
	Else
		JavaWindow("Sistema Integrador v1").JavaTab("Paso 1:").JavaTab("Sección Informativa").JavaTab("Identificación").JavaRadioButton("Balancei_Rdb_No").Set "ON"	
	End If
	
End Function
Function Fun_ITF(itf)

	If itf = "Si" Then
		JavaWindow("Sistema Integrador v1").JavaTab("Paso 1:").JavaTab("Sección Informativa").JavaTab("Identificación").JavaRadioButton("ITF_Rdb_Si").Set "ON"
	Else
		JavaWindow("Sistema Integrador v1").JavaTab("Paso 1:").JavaTab("Sección Informativa").JavaTab("Identificación").JavaRadioButton("ITF_Rdb_No").Set "ON"

	End If
	
End Function
Function Fun_regimen(regimen)
	
	Select Case regimen
		
		Case "General"	
			JavaWindow("Sistema Integrador v1").JavaTab("Paso 1:").JavaTab("Sección Informativa").JavaTab("Identificación").JavaRadioButton("Régimen General").Set "ON"

		Case "Mype"	
			JavaWindow("Sistema Integrador v1").JavaTab("Paso 1:").JavaTab("Sección Informativa").JavaTab("Identificación").JavaRadioButton("Régimen MYPE").Set "ON"

	End Select
	
End Function
Function Fun_MostrarMenu 
	For Iterator = 1 To 15 Step 1
		Set shell = CreateObject("Wscript.Shell")
			shell.SendKeys "{PGUP}"
	Next
End Function




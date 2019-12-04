JavaWindow("Sistema Integrador v1").InsightObject("Pestaña_Reorg").Click

Reorganizacion = "Si"
itan		   = "Si"
anexo_aprobado = "Si"
gasto		   = "Si"
credito		   = "Si"

Call Fun_Reorganizacion(Reorganizacion)
Call Fun_ITAN(itan,anexo_aprobado,gasto,credito)
Call Fun_MostrarMenu()


Function Fun_Reorganizacion(Reorganizacion)
	If Reorganizacion = "Si" Then
		JavaWindow("Sistema Integrador v1").JavaTab("Complete").JavaTab("Sección Informativa").JavaTab("Identificación").JavaRadioButton("Reorg_Si").Set
		JavaWindow("Sistema Integrador v1").JavaTab("Complete").JavaTab("Sección Informativa").JavaTab("Identificación").JavaList("Reorg_233").Select "#2"
		JavaWindow("Sistema Integrador v1").InsightObject("Calendario").Click
		Set shell = CreateObject("Wscript.shell")
			shell.SendKeys"{ENTER}"
		JavaWindow("Sistema Integrador v1").InsightObject("Empresas_parti").Click
		
		For Iterator = 1 To 3 Step 1
			JavaWindow("Sistema Integrador v1").JavaButton("Agregar").Click
			JavaWindow("Sistema Integrador v1").InsightObject("Reorg_ruc").Type 12345876543
			wait 2
			JavaWindow("Sistema Integrador v1").JavaButton("Guardar").Click
			Set shell = CreateObject("Wscript.shell")
				shell.SendKeys"{ENTER}"
		Next
		
		JavaWindow("Sistema Integrador v1").JavaButton("Cancelar").Click

		
	Else
		JavaWindow("Sistema Integrador v1").JavaTab("Complete").JavaTab("Sección Informativa").JavaTab("Identificación").JavaRadioButton("Reorg_No").Set

	End If
End Function
Function Fun_ITAN(itan,anexo_aprobado,gasto,credito)
	
	If itan = "Si" Then
		JavaWindow("Sistema Integrador v1").JavaTab("Complete").JavaTab("Sección Informativa").JavaTab("Identificación").JavaRadioButton("Itan_Si").Set
		
		If anexo_aprobado = "Si" Then
			JavaWindow("Sistema Integrador v1").JavaTab("Complete").JavaTab("Sección Informativa").JavaTab("Identificación").JavaRadioButton("Anexo_Si").Set
			
			If gasto = "Si" Then
				JavaWindow("Sistema Integrador v1").JavaTab("Complete").JavaTab("Sección Informativa").JavaTab("Identificación").JavaCheckBox("Gasto").Set "ON"
			ELSE
				JavaWindow("Sistema Integrador v1").JavaTab("Complete").JavaTab("Sección Informativa").JavaTab("Identificación").JavaCheckBox("Gasto").Set "OFF"
			End If
			
			If credito = "Si" Then
				JavaWindow("Sistema Integrador v1").JavaTab("Complete").JavaTab("Sección Informativa").JavaTab("Identificación").JavaCheckBox("Crédito").Set "ON"
			
			Else 
				JavaWindow("Sistema Integrador v1").JavaTab("Complete").JavaTab("Sección Informativa").JavaTab("Identificación").JavaCheckBox("Crédito").Set "OFF"
				
			End If
			

			

		Else 
			JavaWindow("Sistema Integrador v1").JavaTab("Complete").JavaTab("Sección Informativa").JavaTab("Identificación").JavaRadioButton("Anexo_No").Set

		End If
		
	Else
		
		JavaWindow("Sistema Integrador v1").JavaTab("Complete").JavaTab("Sección Informativa").JavaTab("Identificación").JavaRadioButton("Itan_No").Set

	End If
	
End Function
Function Fun_MostrarMenu 
 	
	For Iterator = 1 To 15 Step 1
		Set shell = CreateObject("Wscript.Shell")
			shell.SendKeys "{PGUP}"
	Next
End Function


		

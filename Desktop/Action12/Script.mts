JavaWindow("Sistema Integrador v1").InsightObject("Pestaña_Regalias").Click

beneficiario = "Si"
Call Fun_Benficiario(beneficiario)

Function Fun_Benficiario(beneficiario)
	
	If beneficiario = "Si" Then
		
		JavaWindow("Sistema Integrador v1").InsightObject("Si").Click
		JavaWindow("Sistema Integrador v1").JavaTab("Paso 1:").JavaTab("Formulario Virtual N°").JavaTab("Información Complementaria").JavaButton("Agregar").Click
		JavaWindow("Sistema Integrador v1").InsightObject("FechaContrato").Click
		Set shell = CreateObject("Wscript.Shell")
				shell.SendKeys "{ENTER}"

		tipo_doc = "PASAPORTE"
		JavaWindow("Sistema Integrador v1").JavaList("TipDoc").Select tipo_doc
		JavaWindow("Sistema Integrador v1").InsightObject("NumDoc").Type 112345623456
		If not tipo_doc = "DNI" or tipo_doc = "RUC" Then
		
			 JavaWindow("Sistema Integrador v1").InsightObject("Beneficiario").type "walther"
		
		End If
		
		JavaWindow("Sistema Integrador v1").InsightObject("Monto").Type 12345
		JavaWindow("Sistema Integrador v1").JavaButton("Guardar").Click

		
		
	End If
	
End Function



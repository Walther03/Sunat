JavaWindow("Sistema Integrador v1").InsightObject("Mas").Click
JavaWindow("Sistema Integrador v1").InsightObject("Pestaña_Desuso").Click


Function Fun_Desuso(desuso)

	If desuso = "Si" Then
		JavaWindow("Sistema Integrador v1").InsightObject("Si").Click

		JavaWindow("Sistema Integrador v1").JavaTab("Paso 1:").JavaTab("Formulario Virtual N°").JavaTab("Información Complementaria").JavaButton("Agregar").Click

		JavaWindow("Sistema Integrador v1").InsightObject("InformeTecnico").Type 123

		JavaWindow("Sistema Integrador v1").InsightObject("FechaInforme").Click
		Set shell = CreateObject("Wscript.Shell")	
				shell.SendKeys "{ENTER}"
				
		JavaWindow("Sistema Integrador v1").JavaList("Sección Informativa").Select "#2"
		JavaWindow("Sistema Integrador v1").InsightObject("NumDoc").Type 44373946

		If not tipo_doc = "DNI" or tipo_doc = "ruc" Then
			JavaWindow("Sistema Integrador v1").InsightObject("Razon").Type "walther"		
		End If
			
		JavaWindow("Sistema Integrador v1").InsightObject("Colegiatura").Type 123
		JavaWindow("Sistema Integrador v1").InsightObject("ValorBien").Type 1234
		JavaWindow("Sistema Integrador v1").InsightObject("Fecha2").Click
		Set shell = CreateObject("Wscript.Shell")
			shell.SendKeys "{ENTER}"
		JavaWindow("Sistema Integrador v1").JavaButton("Guardar").Click
		
		
	Else  
	
		JavaWindow("Sistema Integrador v1").InsightObject("No").Click

	End If
	
End Function



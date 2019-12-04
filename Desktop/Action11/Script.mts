JavaWindow("Sistema Integrador v1").InsightObject("Pestaña_operaciones").Click


han_administrado = "Si"
tercero			 = "Si"
Call Fun_Que_Han_Administrados(han_administrado)
Call Fun_terceros(tercero)

Function Fun_Que_Han_Administrados(han_administrado)
	
	JavaWindow("Sistema Integrador v1").InsightObject("Pestaña que han administrado").Click

	
	If han_administrado = "Si" Then
		JavaWindow("Sistema Integrador v1").InsightObject("Si").Click
		JavaWindow("Sistema Integrador v1").JavaTab("Paso 1:").JavaTab("Sección Informativa").JavaTab("Información Complementaria").JavaTab("Información General").JavaButton("Agregar").Click
		tipo_doc = "PASAPORTE"
		JavaWindow("Sistema Integrador v1").JavaList("TipoDoc").Select tipo_doc
		JavaWindow("Sistema Integrador v1").InsightObject("TipoDoc").type 44374304
		
		If not tipo_doc = "DNI" or tipo_doc = "ruc" Then
		
			JavaWindow("Sistema Integrador v1").InsightObject("Razon").type "walther"
		
		End If
		
		JavaWindow("Sistema Integrador v1").InsightObject("Calendario1").Click
		Set shell = CreateObject("Wscript.Shell")
				shell.SendKeys "{ENTER}"

		indefinido = "Si"
		If indefinido = "Si" Then
			JavaWindow("Sistema Integrador v1").InsightObject("Si").Click
			JavaWindow("Sistema Integrador v1").InsightObject("FinCalendario").Click
			Set shell = CreateObject("Wscript.Shell")
				shell.SendKeys "{ENTER}"
		
			
		Else 
			JavaWindow("Sistema Integrador v1").InsightObject("No").Click
					
		End If
	
		JavaWindow("Sistema Integrador v1").JavaButton("Guardar").Click

	Else 
	
		JavaWindow("Sistema Integrador v1").JavaTab("Paso 1:").JavaTab("Sección Informativa").JavaTab("Información Complementaria").JavaTab("Información General").JavaRadioButton("indeterminado_No").Set

		
	End If
	
	JavaWindow("Sistema Integrador v1").JavaButton("Cancelar").Click

	
End Function

Function Fun_terceros(tercero)
	
	JavaWindow("Sistema Integrador v1").InsightObject("Pestaña_tercero").Click

	
	If tercero = "Si" Then
		JavaWindow("Sistema Integrador v1").InsightObject("Si").Click
		JavaWindow("Sistema Integrador v1").JavaTab("Paso 1:").JavaTab("Sección Informativa").JavaTab("Información Complementaria").JavaTab("Información General").JavaButton("Agregar").Click
		tipo_doc = "PASAPORTE"
		JavaWindow("Sistema Integrador v1").JavaList("TipoDoc").Select tipo_doc
		JavaWindow("Sistema Integrador v1").InsightObject("TipoDoc").type 44374304
		
		If not tipo_doc = "DNI" or tipo_doc = "ruc" Then
		
			JavaWindow("Sistema Integrador v1").InsightObject("Razon").type "waalther"
		
		End If
		
		JavaWindow("Sistema Integrador v1").InsightObject("Calendario1").Click
		Set shell = CreateObject("Wscript.Shell")
				shell.SendKeys "{ENTER}"

		indefinido = "Si"
		If indefinido = "Si" Then
			JavaWindow("Sistema Integrador v1").InsightObject("Si").Click
			JavaWindow("Sistema Integrador v1").InsightObject("FinCalendario").Click
			Set shell = CreateObject("Wscript.Shell")
				shell.SendKeys "{ENTER}"
		
			
		Else 
			JavaWindow("Sistema Integrador v1").InsightObject("No").Click
					
		End If
	
		JavaWindow("Sistema Integrador v1").JavaButton("Guardar").Click

	Else 
	
		JavaWindow("Sistema Integrador v1").JavaTab("Paso 1:").JavaTab("Sección Informativa").JavaTab("Información Complementaria").JavaTab("Información General").JavaRadioButton("indeterminado_No").Set

		
	End If
	
	JavaWindow("Sistema Integrador v1").JavaButton("Cancelar").Click

	
End Function


		

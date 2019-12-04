JavaWindow("Sistema Integrador v1").InsightObject("Pestaña_mermas").Click

merma 	  = "Si"
desmedros = "Si"
call Fun_Merma(merma)
call Fun_Desmedros(desmedros)

Function Fun_Merma(merma)

	If merma = "Si" Then
		JavaWindow("Sistema Integrador v1").InsightObject("Merma").Click
		JavaWindow("Sistema Integrador v1").InsightObject("Si").Click
	
		JavaWindow("Sistema Integrador v1").JavaTab("Complete").JavaTab("Sección Informativa").JavaTab("Información Complementaria").JavaTab("Pérdidas extraordinarias").JavaButton("Agregar").Click
	
		JavaWindow("Sistema Integrador v1").JavaList("Registro de Detalle de").Select "#1"
		JavaWindow("Sistema Integrador v1").JavaEdit("Informe").Set 123
		JavaWindow("Sistema Integrador v1").InsightObject("Calendario").Click
		wait 2
		Set shell = CreateObject("Wscript.Shell")
				shell.SendKeys "{ENTER}"
		wait 1
		
		tipodoc = "DNI"
		JavaWindow("Sistema Integrador v1").InsightObject("TipDoc").Click
		Select Case tipodoc
			Case "DNI"
				JavaWindow("Sistema Integrador v1").InsightObject("InsightObject_3").Click
			 @@ hightlight id_;_3_;_script infofile_;_ZIP::ssf11.xml_;_
			Case "Otros"	 @@ hightlight id_;_7_;_script infofile_;_ZIP::ssf12.xml_;_
				JavaWindow("Sistema Integrador v1").InsightObject("OtrosTipos").Click
				
			Case "ruc"	
				JavaWindow("Sistema Integrador v1").InsightObject("ruc").Click
				
			Case "carnet"	
				JavaWindow("Sistema Integrador v1").InsightObject("Carnet").Click
			
			Case "pasaporte"	
				JavaWindow("Sistema Integrador v1").InsightObject("Pasaporte").Click
			 @@ hightlight id_;_34_;_script infofile_;_ZIP::ssf10.xml_;_
			
		End Select
		JavaWindow("Sistema Integrador v1").InsightObject("NumDoc").Click
		JavaWindow("Sistema Integrador v1").InsightObject("NumDoc").type 44393796
		If not tipodoc = "DNI" or  tipodoc = "ruc" Then
			
			JavaWindow("Sistema Integrador v1").InsightObject("Razon").Type "razon"

		End If
		JavaWindow("Sistema Integrador v1").InsightObject("Colegiatura").type 123
		JavaWindow("Sistema Integrador v1").InsightObject("ImporteMerma").type 123

		
		JavaWindow("Sistema Integrador v1").JavaButton("Guardar").Click

	Else 
	
		JavaWindow("Sistema Integrador v1").JavaTab("Complete").JavaTab("Sección Informativa").JavaTab("Información Complementaria").JavaTab("Pérdidas extraordinarias").JavaRadioButton("No").Set

	End If
	
	
	
End Function

Function Fun_Desmedros(desmedros)
	
	If desmedros = "Si" Then
		JavaWindow("Sistema Integrador v1").InsightObject("Pestaña_Demedros").Click
		JavaWindow("Sistema Integrador v1").InsightObject("Si").Click
	
		JavaWindow("Sistema Integrador v1").JavaTab("Complete").JavaTab("Sección Informativa").JavaTab("Información Complementaria").JavaTab("Pérdidas extraordinarias").JavaButton("Agregar").Click
		
		desmedro_ante = "Notario"
		Select Case desmedro_ante
			Case "Notario"
				JavaWindow("Sistema Integrador v1").JavaList("Registro de Detalle de_2").Select "#1"
				JavaWindow("Sistema Integrador v1").InsightObject("NumDoc").type 10282298738

				JavaWindow("Sistema Integrador v1").InsightObject("desmedro_Calendario1").Click
				Set shell = CreateObject("Wscript.shell")
					shell.SendKeys("{UP}")
					shell.SendKeys("{UP}")
					shell.SendKeys("{ENTER}")				
					
				JavaWindow("Sistema Integrador v1").InsightObject("desmedro_Calendario2").Click
				Set shell = CreateObject("Wscript.shell")
					shell.SendKeys("{ENTER}")	
				
				JavaWindow("Sistema Integrador v1").InsightObject("ImporteDesmedro").type 123
				JavaWindow("Sistema Integrador v1").JavaButton("Guardar").Click


			Case "Juez"
				JavaWindow("Sistema Integrador v1").JavaList("Registro de Detalle de_2").Select "#2"
				JavaWindow("Sistema Integrador v1").InsightObject("IdentificacionJueza").type 123
				
					JavaWindow("Sistema Integrador v1").InsightObject("desmedro_Calendario1").Click
				Set shell = CreateObject("Wscript.shell")
					shell.SendKeys("{UP}")
					shell.SendKeys("{UP}")
					shell.SendKeys("{ENTER}")				
					
				JavaWindow("Sistema Integrador v1").InsightObject("desmedro_Calendario2").Click
				Set shell = CreateObject("Wscript.shell")
					shell.SendKeys("{ENTER}")	
				
				JavaWindow("Sistema Integrador v1").InsightObject("ImporteDesmedro").type 123
				JavaWindow("Sistema Integrador v1").JavaButton("Guardar").Click

		End Select
		
	Else 
	
		JavaWindow("Sistema Integrador v1").JavaTab("Complete").JavaTab("Sección Informativa").JavaTab("Información Complementaria").JavaTab("Pérdidas extraordinarias").JavaRadioButton("No").Set

		
	End If
	
End Function


		

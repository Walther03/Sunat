JavaWindow("Sistema Integrador v1").InsightObject("Pestaña_Exoneracion").Click


exoneracion 		= "Si"
inafecta 			= "No"
donacion 			= "Si"
convenio 			= "Si"
moneda_extranjera 	= "Si"
tipo_moneda			= "Nacional"
beneficio 			= "Si"

Call Fun_Exoneracion(exoneracion)
If exoneracion = "No" Then
	Call Fun_Inafecta(exoneracion)
	
	If inafecta = "No" Then
		Call Fun_Donacion(donacion)
		Call Fun_Convenio(convenio,moneda_extranjera,tipo_moneda)
		Call Fun_Beneficios(beneficio)
	End If
	
End If


Function Fun_Exoneracion(exoneracion)
	
	If exoneracion = "Si" Then
		
		JavaWindow("Sistema Integrador v1").JavaTab("Complete").JavaTab("Paso 1:").JavaTab("Formulario Virtual N°").JavaRadioButton("Exoneracion_Si").Set
		JavaWindow("Sistema Integrador v1").JavaTab("Complete").JavaTab("Paso 1:").JavaTab("Formulario Virtual N°").JavaList("210").Select "#5"
		JavaWindow("Sistema Integrador v1").JavaTab("Complete").JavaTab("Paso 1:").JavaTab("Formulario Virtual N°").JavaEdit("216").Set "Especifique"
		
	Else 

		JavaWindow("Sistema Integrador v1").JavaTab("Complete").JavaTab("Paso 1:").JavaTab("Formulario Virtual N°").JavaRadioButton("Exoneracion_No").Set

	End If
	
End Function
Function Fun_Inafecta(inafecta)
	
	If inafecta = "Si" Then
	
	JavaWindow("Sistema Integrador v1").InsightObject("Inafecta_Si").Click
	JavaWindow("Mensaje").InsightObject("Mensaje_btn_Si").Click
	JavaWindow("Sistema Integrador v1").JavaTab("Complete").JavaTab("Paso 1:").JavaTab("Formulario Virtual N°").JavaList("221").Select "#6"
	JavaWindow("Sistema Integrador v1").JavaTab("Complete").JavaTab("Paso 1:").JavaTab("Formulario Virtual N°").JavaEdit("222").Set "Especifique"

	Else 

	JavaWindow("Sistema Integrador v1").JavaTab("Complete").JavaTab("Paso 1:").JavaTab("Formulario Virtual N°").JavaRadioButton("Inafecta_No").Set
	
	End If
		
	
	
End Function
Function Fun_Donacion(donacion)
	
	If donacion = "Si" Then
		JavaWindow("Sistema Integrador v1").JavaTab("Complete").JavaTab("Paso 1:").JavaTab("Formulario Virtual N°").JavaRadioButton("Don_Si").Set
		JavaWindow("Sistema Integrador v1").InsightObject("Don_Txt227").Click
		
		JavaWindow("Sistema Integrador v1").JavaButton("Don_Agregar").Click
		JavaWindow("Sistema Integrador v1").JavaList("Don_Registro de Detalle de").Select "#0"
		JavaWindow("Sistema Integrador v1").JavaList("Don_¿Está exonerado totalmente").Select "#2"
		JavaWindow("Sistema Integrador v1").JavaEdit("Don_Tipo de Documento del").Set 10404266417
		JavaWindow("Sistema Integrador v1").InsightObject("Don_Calendario").Click
		
		Set shell = CreateObject("Wscript.Shell")
			shell.SendKeys "{ENTER}"
		
		wait 2
		JavaWindow("Sistema Integrador v1").JavaEdit("Don_Tabla sin contenido").Set 20
		JavaWindow("Sistema Integrador v1").JavaButton("Don_Guardar").Click
		
		JavaWindow("Sistema Integrador v1").JavaButton("Don_Cancelar").Click

	
	Else 
		
		JavaWindow("Sistema Integrador v1").JavaTab("Complete").JavaTab("Paso 1:").JavaTab("Formulario Virtual N°").JavaRadioButton("Don_No").Set


	End If

	
End Function
Function Fun_Convenio(convenio,moneda_extranjera,tipo_moneda)
	
	If convenio = "Si" Then
		JavaWindow("Sistema Integrador v1").JavaTab("Complete").JavaTab("Paso 1:").JavaTab("Formulario Virtual N°").JavaRadioButton("Conv_Si").Set
		JavaWindow("Sistema Integrador v1").JavaTab("Complete").JavaTab("Paso 1:").JavaTab("Formulario Virtual N°").JavaEdit("Conv_224").Set 123

		If moneda_extranjera = "Si" Then
		
			JavaWindow("Sistema Integrador v1").JavaTab("Complete").JavaTab("Paso 1:").JavaTab("Formulario Virtual N°").JavaRadioButton("Mon_Si").Set
			
			Select Case tipo_moneda
			
				Case "Nacional"
					JavaWindow("Sistema Integrador v1").JavaTab("Complete").JavaTab("Paso 1:").JavaTab("Formulario Virtual N°").JavaRadioButton("Moneda Nacional").Set

				Case "Extranjera"
					JavaWindow("Sistema Integrador v1").JavaTab("Complete").JavaTab("Paso 1:").JavaTab("Formulario Virtual N°").JavaRadioButton("Moneda Extranjera").Set

			End Select
			
		Else 
			
			JavaWindow("Sistema Integrador v1").JavaTab("Complete").JavaTab("Paso 1:").JavaTab("Formulario Virtual N°").JavaRadioButton("Mon_No").Set

		End If

	Else 	
	
		JavaWindow("Sistema Integrador v1").JavaTab("Complete").JavaTab("Paso 1:").JavaTab("Formulario Virtual N°").JavaRadioButton("Conv_No").Set

	End If
	
End Function

Function Fun_Beneficios(beneficio)
	
	If beneficio = "Si" Then
		
		JavaWindow("Sistema Integrador v1").JavaTab("Complete").JavaTab("Paso 1:").JavaTab("Formulario Virtual N°").JavaRadioButton("Ben_Si").Set
		JavaWindow("Sistema Integrador v1").JavaTab("Complete").JavaTab("Paso 1:").JavaTab("Formulario Virtual N°").JavaList("Ben_Cmb_199").Select "#3"
		JavaWindow("Sistema Integrador v1").JavaTab("Complete").JavaTab("Paso 1:").JavaTab("Formulario Virtual N°").JavaEdit("Ben_201").Set "Especifique"

	Else
		JavaWindow("Sistema Integrador v1").JavaTab("Complete").JavaTab("Paso 1:").JavaTab("Formulario Virtual N°").JavaRadioButton("Ben_No").Set

	End If
	
End Function
		
		

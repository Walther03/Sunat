JavaWindow("Sistema Integrador v1").InsightObject("Mas").Click @@ hightlight id_;_1_;_script infofile_;_ZIP::ssf1.xml_;_
JavaWindow("Sistema Integrador v1").InsightObject("Pestaña_GastosInvestigaacion").Click @@ hightlight id_;_6_;_script infofile_;_ZIP::ssf2.xml_;_



gatosinvestigacion = "Si"
beneficio = "Con Beneficio"
proyecto  = "Directamente"

Call Fun_GastosInvestigacion(gatosinvestigacion,beneficio , proyecto)
Call Fun_MostrarMenu()
Function Fun_GastosInvestigacion(gatosinvestigacion,beneficio , proyecto)
		
		
	If gatosinvestigacion = "Si" Then
		JavaWindow("Sistema Integrador v1").InsightObject("Si").Click
		Select Case beneficio
			Case "Sin beneficio"
				JavaWindow("Sistema Integrador v1").InsightObject("Pestaña_Sinbeneficio").Click
				JavaWindow("Sistema Integrador v1").JavaTab("Paso 1:").JavaTab("Formulario Virtual N°").JavaTab("Información Complementaria").JavaTab("¿Ha realizado gastos en").JavaButton("Agregar").Click
				JavaWindow("Sistema Integrador v1").JavaRadioButton("No vinculado al Giro de").Set "ON"
				JavaWindow("Sistema Integrador v1").JavaRadioButton("Vinculado al Giro de Negocio").Set "ON"
				JavaWindow("Sistema Integrador v1").InsightObject("DescripcionProyecto").Type "Descripcion"
				JavaWindow("Sistema Integrador v1").JavaEdit("ImporteTotal").Set 123
				JavaWindow("Sistema Integrador v1").JavaEdit("MontoDeducible").Set 123
				JavaWindow("Sistema Integrador v1").InsightObject("NoDeducible").Type 123
				JavaWindow("Sistema Integrador v1").JavaButton("Guardar").Click
		
			Case "Con Beneficio"	
				JavaWindow("Sistema Integrador v1").InsightObject("Pestaña_Conbeneficio").Click
				JavaWindow("Sistema Integrador v1").JavaTab("Paso 1:").JavaTab("Formulario Virtual N°").JavaTab("Información Complementaria").JavaTab("¿Ha realizado gastos en").JavaButton("Agregar").Click
			
				Select Case proyecto
					Case "Directamente"
						JavaWindow("Sistema Integrador v1").InsightObject("Proyecto").Click
						JavaWindow("Sistema Integrador v1").JavaRadioButton("No vinculado al Giro de").Set
						JavaWindow("Sistema Integrador v1").JavaRadioButton("Vinculado al Giro de Negocio").Set
						JavaWindow("Sistema Integrador v1").InsightObject("DescripcionProyecto").type "Descripcion"
						JavaWindow("Sistema Integrador v1").InsightObject("FechaCalificacion").Click
						Set shell = CreateObject("Wscript.shell")
							shell.SendKeys "{UP}"
							shell.SendKeys "{ENTER}"
					
						JavaWindow("Sistema Integrador v1").InsightObject("FechaAutorizacion").Click
						Set shell = CreateObject("Wscript.shell")
							shell.SendKeys "{ENTER}"
					
						JavaWindow("Sistema Integrador v1").JavaEdit("Importe Total").Set 123
					
						JavaWindow("Sistema Integrador v1").JavaEdit("Monto Deducible").Set 123
						JavaWindow("Sistema Integrador v1").InsightObject("MontoNO2").Type 123
						JavaWindow("Sistema Integrador v1").JavaList("Beneficio Ley 30309").Select "#1"
						JavaWindow("Sistema Integrador v1").JavaEdit("Beneficio Ley 30309").Set 123
						JavaWindow("Sistema Integrador v1").JavaButton("Guardar").Click

					Case "Proyecto"
						JavaWindow("Sistema Integrador v1").InsightObject("AtravezCentro").Click				
						JavaWindow("Sistema Integrador v1").JavaRadioButton("Proyecto_Si").Set
						JavaWindow("Sistema Integrador v1").JavaRadioButton("Proyecto_No").Set

						JavaWindow("Sistema Integrador v1").JavaButton("Guardar").Click	
				End Select
			
		End Select
		
	JavaWindow("Sistema Integrador v1").JavaTab("Paso 1:").JavaTab("Formulario Virtual N°").JavaTab("Información Complementaria").JavaRadioButton("Si").Set

	Else

		
	
	
	End If
		
	
End Function


Function Fun_MostrarMenu()
	For Iterator = 1 To 15 Step 1
		Set shell = CreateObject("Wscript.Shell")
			shell.SendKeys "{PGUP}"
	Next
End Function

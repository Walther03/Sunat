
JavaWindow("Sistema Integrador v1").InsightObject("Mas").Click @@ hightlight id_;_1_;_script infofile_;_ZIP::ssf1.xml_;_
JavaWindow("Sistema Integrador v1").InsightObject("Pestaña_RentaInternacional").Click @@ hightlight id_;_6_;_script infofile_;_ZIP::ssf2.xml_;_

Function Fun_RentaInternacional(renta_internacional)
	
	If renta_internacional = "Si" Then
		
		JavaWindow("Sistema Integrador v1").InsightObject("Si").Click
		JavaWindow("Sistema Integrador v1").JavaTab("Gastos realizados con").JavaTab("Gastos realizados con").JavaTab("Gastos realizados con").JavaButton("Agregar").Click
		JavaWindow("Sistema Integrador v1").JavaList("TipoRenta").Select "#2"
		JavaWindow("Sistema Integrador v1").InsightObject("Pais").Click
		Set shell = CreateObject("Wscript.shell")
			shell.SendKeys "{DOWN}"
			shell.SendKeys "{ENTER}"
	
		JavaWindow("Sistema Integrador v1").InsightObject("Aplica").Click
		Set shell = CreateObject("Wscript.shell")
			shell.SendKeys "{DOWN}"
			shell.SendKeys "{DOWN}"
			shell.SendKeys "{ENTER}"
	
		JavaWindow("Sistema Integrador v1").InsightObject("Monto").Type 123

		JavaWindow("Sistema Integrador v1").JavaButton("Aceptar").Click

	Else
	
		JavaWindow("Sistema Integrador v1").JavaTab("Gastos realizados con").JavaTab("Gastos realizados con").JavaTab("Gastos realizados con").JavaRadioButton("No").Set

		
	End If
	
End Function


JavaWindow("Sistema Integrador v1").InsightObject("Mas").Click @@ hightlight id_;_10_;_script infofile_;_ZIP::ssf5.xml_;_
JavaWindow("Sistema Integrador v1").InsightObject("Pestaña_perdida").Click @@ hightlight id_;_18_;_script infofile_;_ZIP::ssf7.xml_;_

inutil = "No"
delictuoso = "Si"
motivo= "Fortuito"
monto = 122	

Call Fun_perdidas(inutil, delictuoso , motivo , monto)

Function Fun_perdidas(inutil, delictuoso , motivo , monto)


	JavaWindow("Sistema Integrador v1").InsightObject("Si").Click
	JavaWindow("Sistema Integrador v1").JavaTab("Alquileres Pagados").JavaTab("Alquileres Pagados").JavaTab("Alquileres Pagados").JavaButton("Agregar").Click
	JavaWindow("Sistema Integrador v1").InsightObject("FechaPerdida").Click
	Set shell = CreateObject("Wscript.shell")
		shell.SendKeys "{ENTER}"

	If perdida_cubierta = "Si" Then
		JavaWindow("Sistema Integrador v1").JavaRadioButton("Cubierta_Si").Set

	Else 
		JavaWindow("Sistema Integrador v1").JavaRadioButton("Cubierta_No").Set
	
	End If


	If inutil = "Si" Then
		JavaWindow("Sistema Integrador v1").JavaRadioButton("Inutil_Si").Set

	Else 
		JavaWindow("Sistema Integrador v1").JavaRadioButton("Inutil_No").Set
	
		If delictuoso = "Si" Then
			JavaWindow("Sistema Integrador v1").InsightObject("delictuoso_si").Click

			JavaWindow("Sistema Integrador v1").InsightObject("Delictuoso_fecha").Click
			Set shell = CreateObject("Wscript.shell")
				shell.SendKeys "{UP}"
				shell.SendKeys "{UP}"
				shell.SendKeys "{ENTER}"
				
			JavaWindow("Sistema Integrador v1").JavaList("Identificación del Juzgado:").Select "#2"

		
		Else 
			JavaWindow("Sistema Integrador v1").JavaRadioButton("Delictuoso_No").Set

		End If
	
	End If
	

	Select Case motivo
		
		Case "Fortuito"
			JavaWindow("Sistema Integrador v1").JavaEdit("Por Caso Fortuito o").Set monto
		
		Case "Delito"
			JavaWindow("Sistema Integrador v1").JavaEdit("Por delitos cometidos").Set monto

		
	End Select
	

JavaWindow("Sistema Integrador v1").JavaButton("Guardar").Click

	
	
	
	
End Function


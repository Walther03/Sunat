JavaWindow("Sistema Integrador v1").InsightObject("Pestaña_Mineria").Click

mineria		 			= "No"
novinculadas 			= 100
vinculadas			 	= 100
Hidrocarburo			= "Si"
exploracion				= 100
actividad_relacionada	= 100
otras					= 100

Call Fun_Mineria( mineria , novinculadas , vinculadas )
Call Fun_MostrarMenu()


Function Fun_Mineria( mineria , novinculadas , vinculadas )
	
	If mineria = "Si" Then
	
		JavaWindow("Sistema Integrador v1").JavaTab("Complete").JavaTab("Sección Informativa").JavaTab("Identificación").JavaRadioButton("Mineria_Si").Set
		JavaWindow("Sistema Integrador v1").JavaTab("Complete").JavaTab("Sección Informativa").JavaTab("Identificación").JavaEdit("Min_228").Set novinculadas
		JavaWindow("Sistema Integrador v1").JavaTab("Complete").JavaTab("Sección Informativa").JavaTab("Identificación").JavaEdit("Min_238").Set vinculadas

	Else
	
		JavaWindow("Sistema Integrador v1").JavaTab("Complete").JavaTab("Sección Informativa").JavaTab("Identificación").JavaRadioButton("Mineria_No").Set
		Call Fun_Hidrocarburo(Hidrocarburo , exploracion , actividad_relacionada , otras)
		
	End If
End Function

Function Fun_Hidrocarburo(Hidrocarburo , exploracion , actividad_relacionada , otras)
	
	If Hidrocarburo = "Si" Then
	
		JavaWindow("Sistema Integrador v1").JavaTab("Complete").JavaTab("Sección Informativa").JavaTab("Identificación").JavaRadioButton("Hidrocar_Si").Set
		
		JavaWindow("Sistema Integrador v1").JavaTab("Complete").JavaTab("Sección Informativa").JavaTab("Identificación").JavaEdit("Hidrocar_275").Set exploracion
		JavaWindow("Sistema Integrador v1").JavaTab("Complete").JavaTab("Sección Informativa").JavaTab("Identificación").JavaEdit("Hidrocar_276").Set actividad_relacionada
		JavaWindow("Sistema Integrador v1").JavaTab("Complete").JavaTab("Sección Informativa").JavaTab("Identificación").JavaEdit("Hidrocar_277").Set otras

		JavaWindow("Sistema Integrador v1").JavaTab("Complete").JavaTab("Sección Informativa").JavaTab("Identificación").JavaRadioButton("Hidrocar_Si").Set
	Else
	
		JavaWindow("Sistema Integrador v1").JavaTab("Complete").JavaTab("Sección Informativa").JavaTab("Identificación").JavaRadioButton("Hidrocar_No").Set

	End If
		
End Function

Function Fun_MostrarMenu()
	
	For Iterator = 1 To 15 Step 1
		Set shell = CreateObject("Wscript.Shell")
			shell.SendKeys "{PGUP}"
	Next
End Function

JavaWindow("Sistema Integrador v1").InsightObject("Informacion Complementaria").Click


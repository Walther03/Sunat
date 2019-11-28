JavaWindow("Sistema Integrador v1").InsightObject("Pestaña_Contrato").Click

Dim contrato , AgregarContrato ,Des_contrato , num_participantes 
contrato 			= "Si"
AgregarContrato 	= "Si"
Des_contrato 		= "Descripcion"
num_participantes	=  2


Call Fun_Contrato(contrato)
Call Fun_Agregar(AgregarContrato)
Call Fun_NuevoParticipador(num_participantes)


Function Fun_Contrato(contrato)
	
	If contrato = "Si" Then
		JavaWindow("Sistema Integrador v1").InsightObject("Btn_Si").Click
			
	Else 
		JavaWindow("Sistema Integrador v1").JavaTab("Paso 1:").JavaTab("Sección Informativa").JavaTab("Identificación").JavaRadioButton("No").Set "ON"

	End If
	
End Function

Function Fun_Agregar(AgregarContrato)

	If AgregarContrato = "Si" Then
	
		JavaWindow("Sistema Integrador v1").JavaTab("Paso 1:").JavaTab("Sección Informativa").JavaTab("Identificación").JavaButton("Agregar").Click
		JavaWindow("Sistema Integrador v1").InsightObject("Desc").Type Des_contrato
		
	End If
		
End Function

 Function Fun_NuevoParticipador(num_participantes) 
 
 	For Iterator = 1 To 100 Step 1
 	
 		tipodoc			= "DNI"
 		numdoc 			= 44394190
 		participacion	= 0.001
 		
 		JavaWindow("Sistema Integrador v1").JavaButton("Nuevo participante").Click
	
		JavaWindow("Sistema Integrador v1").JavaList("Detalle de Contratos de").Select tipodoc
		JavaWindow("Sistema Integrador v1").InsightObject("DocNum").Type numdoc
		Set shell = CreateObject("Wscript.Shell")
			shell.SendKeys "{ENTER}"
		WAIT 2
		JavaWindow("Sistema Integrador v1").InsightObject("%").Type participacion

		JavaWindow("Sistema Integrador v1").JavaRadioButton("Como Operador").Set "ON"
'		JavaWindow("Sistema Integrador v1").JavaRadioButton("Como Participe").Set
		JavaWindow("Sistema Integrador v1").JavaButton("Guardar").Click
		Set shell = CreateObject("Wscript.Shell")
			shell.SendKeys "{ENTER}"
		WAIT 2
		
 	Next
 	JavaWindow("Sistema Integrador v1").JavaButton("Cancelar").Click
 End Function
 

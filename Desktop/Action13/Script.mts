 @@ hightlight id_;_9_;_script infofile_;_ZIP::ssf2.xml_;_
 
JavaWindow("Sistema Integrador v1").InsightObject("Mas").Click
 @@ hightlight id_;_1_;_script infofile_;_ZIP::ssf1.xml_;_
JavaWindow("Sistema Integrador v1").InsightObject("InfomacionGeneral").Click
perdida = "Si"
arrastre = "A"
Call DatosContador()
Call Instrumentos(perdida,arrastre)
Call RepresanteLegal()

Function DatosContador()
	tipo_doc = "DOC. NACIONAL DE IDENTIDAD"
	JavaWindow("Sistema Integrador v1").JavaTab("Paso 1:").JavaTab("Formulario Virtual N°").JavaTab("Información Complementaria").JavaList("TipoDoc").Select tipo_doc

	Select Case tipo_doc
		Case "DOC. NACIONAL DE IDENTIDAD"
			JavaWindow("Sistema Integrador v1").JavaTab("Paso 1:").JavaTab("Formulario Virtual N°").JavaTab("Información Complementaria").JavaEdit("DNI").Set 44377871
			JavaWindow("Sistema Integrador v1").JavaTab("Paso 1:").JavaTab("Formulario Virtual N°").JavaTab("Información Complementaria").JavaEdit("CPC").Set 123
			JavaWindow("Sistema Integrador v1").JavaTab("Paso 1:").JavaTab("Formulario Virtual N°").JavaTab("Información Complementaria").JavaEdit("Correo Electrónico 1").Set "automatizacion@gmail.com"
			JavaWindow("Sistema Integrador v1").JavaTab("Paso 1:").JavaTab("Formulario Virtual N°").JavaTab("Información Complementaria").JavaEdit("Correo Electrónico 2").Set "automatizacion@gmail.com"
			JavaWindow("Sistema Integrador v1").JavaTab("Paso 1:").JavaTab("Formulario Virtual N°").JavaTab("Información Complementaria").JavaEdit("Teléfono Fijo").Set 1234567
			JavaWindow("Sistema Integrador v1").JavaTab("Paso 1:").JavaTab("Formulario Virtual N°").JavaTab("Información Complementaria").JavaEdit("Teléfono Celular").Set 957234640
		
		Case "REG. UNICO DE CONTRIBUYENTES"
			JavaWindow("Sistema Integrador v1").JavaTab("Paso 1:").JavaTab("Formulario Virtual N°").JavaTab("Información Complementaria").JavaEdit("RUC").Set 10077888492
			JavaWindow("Sistema Integrador v1").JavaTab("Paso 1:").JavaTab("Formulario Virtual N°").JavaTab("Información Complementaria").JavaEdit("CPC").Set 123
			JavaWindow("Sistema Integrador v1").JavaTab("Paso 1:").JavaTab("Formulario Virtual N°").JavaTab("Información Complementaria").JavaEdit("Correo Electrónico 1").Set "automatizacion@gmail.com"
			JavaWindow("Sistema Integrador v1").JavaTab("Paso 1:").JavaTab("Formulario Virtual N°").JavaTab("Información Complementaria").JavaEdit("Correo Electrónico 2").Set "automatizacion@gmail.com"
			JavaWindow("Sistema Integrador v1").JavaTab("Paso 1:").JavaTab("Formulario Virtual N°").JavaTab("Información Complementaria").JavaEdit("Teléfono Fijo").Set  1234567
			JavaWindow("Sistema Integrador v1").JavaTab("Paso 1:").JavaTab("Formulario Virtual N°").JavaTab("Información Complementaria").JavaEdit("Teléfono Celular").Set 957234640
		
	
	End Select


End Function


Function Instrumentos(perdida,arrastre)

	If perdida = "Si" Then
		JavaWindow("Sistema Integrador v1").JavaTab("Paso 1:").JavaTab("Formulario Virtual N°").JavaTab("Información Complementaria").JavaRadioButton("Perdida_Si").Set

	Else 
		JavaWindow("Sistema Integrador v1").JavaTab("Paso 1:").JavaTab("Formulario Virtual N°").JavaTab("Información Complementaria").JavaRadioButton("Perdida_No").Set

	End If
	
	If arrastre = "A" Then
		JavaWindow("Sistema Integrador v1").JavaTab("Paso 1:").JavaTab("Formulario Virtual N°").JavaTab("Información Complementaria").JavaRadioButton("Perdida_A").Set

	Else 
		JavaWindow("Sistema Integrador v1").JavaTab("Paso 1:").JavaTab("Formulario Virtual N°").JavaTab("Información Complementaria").JavaRadioButton("Perdida_B").Set


	End If
	
End Function

Function RepresanteLegal()
	
	JavaWindow("Sistema Integrador v1").JavaTab("Paso 1:").JavaTab("Formulario Virtual N°").JavaTab("Información Complementaria").JavaList("Repre_TipoDoc").Select "#3"
	JavaWindow("Sistema Integrador v1").JavaTab("Paso 1:").JavaTab("Formulario Virtual N°").JavaTab("Información Complementaria").JavaEdit("Repre_TipoDoc").Set 123456765432

	
End Function

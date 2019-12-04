JavaWindow("Sistema Integrador v1").InsightObject("Pestaña_mineria").Click @@ hightlight id_;_2_;_script infofile_;_ZIP::ssf1.xml_;_
JavaWindow("Sistema Integrador v1").JavaTab("Formulario Virtual N°").JavaTab("Sección Determinativa").JavaTab("Estados Financieros").JavaButton("Agregar").Click
	JavaWindow("Sistema Integrador v1").JavaEdit("NroCodigo").Set 123
	JavaWindow("Sistema Integrador v1").JavaEdit("NombreInversion").Set "Nombre"
	JavaWindow("Sistema Integrador v1").JavaEdit("TotalActivo").Set 123
	JavaWindow("Sistema Integrador v1").JavaEdit("TotalPasivo").Set 123
	JavaWindow("Sistema Integrador v1").JavaEdit("TotalPatrimonio").Set 123
	JavaWindow("Sistema Integrador v1").JavaEdit("VentaNetas").Set 123
	JavaWindow("Sistema Integrador v1").JavaEdit("CostoVentas").Set 123
	JavaWindow("Sistema Integrador v1").JavaEdit("OtrosIngresos").Set 123
	JavaWindow("Sistema Integrador v1").JavaEdit("EjercicioPositivo").Set 123
	JavaWindow("Sistema Integrador v1").JavaEdit("EjercicioNegativo").Set 123
	
	JavaWindow("Sistema Integrador v1").InsightObject("OtrosGastos").Click
		JavaWindow("Sistema Integrador v1").JavaEdit("SustanciasMineras").Set 123
		JavaWindow("Sistema Integrador v1").JavaEdit("GastosAtribuidosDirectamente").Set 123
		JavaWindow("Sistema Integrador v1").JavaEdit("GastosAtribuidosIndirectamente").Set 123
		JavaWindow("Sistema Integrador v1").JavaButton("Aceptar").Click
	
	JavaWindow("Sistema Integrador v1").InsightObject("DistribucionLegal").Click
		JavaWindow("Sistema Integrador v1").JavaEdit("ParticipacionComunidad").Set 123
		JavaWindow("Sistema Integrador v1").JavaEdit("ParticipacionPatrimonial").Set 123
		JavaWindow("Sistema Integrador v1").JavaEdit("ParticipacionTrabajadores").Set 123
		JavaWindow("Sistema Integrador v1").InsightObject("InvestigacionCientifica").Type 123
		JavaWindow("Sistema Integrador v1").JavaEdit("DistribucionLegal").Set 123
		JavaWindow("Sistema Integrador v1").JavaButton("Aceptar").Click
	
	JavaWindow("Sistema Integrador v1").JavaEdit("PerdidaOtras").Set 123
	JavaWindow("Sistema Integrador v1").InsightObject("Adiciones").Click
		JavaWindow("Sistema Integrador v1").JavaEdit("Monto Adición (*)").Set 1
		Set shell = CreateObject("Wscript.shell")
			shell.SendKeys "{TAB}"
		
		For Iterator = 1 To 32 Step 1
			JavaWindow("Sistema Integrador v1").InsightObject("Cuadrado").Type Iterator + 1
			Set shell = CreateObject("Wscript.shell")
			shell.SendKeys "{TAB}"
			
		Next
		shell.SendKeys "{ENTER}"
		
		JavaWindow("Sistema Integrador v1").JavaEdit("34").Set 34
			Set shell = CreateObject("Wscript.shell")
			shell.SendKeys "{TAB}"
			WAIT 1
			JavaWindow("Sistema Integrador v1").InsightObject("Cuadrado2").Type 35
			
			wait 2
			Cont = 36
		For Iterator = 1 To 12 Step 1
			Cont = cont + 1
			JavaWindow("Sistema Integrador v1").InsightObject("Cuadrado3").Type Cont
			Set shell = CreateObject("Wscript.shell")
			shell.SendKeys "{TAB}"
		Next
		

	
	
	JavaWindow("Sistema Integrador v1").JavaEdit("PerdidaEjercicios").Set 123
	JavaWindow("Sistema Integrador v1").JavaEdit("Tasa").Set 123
	JavaWindow("Sistema Integrador v1").JavaEdit("ImporteMinimo").Set 123
	


JavaWindow("Sistema Integrador v1").InsightObject("Pestaña_PasivoPatrimonio").Click @@ hightlight id_;_1_;_script infofile_;_ZIP::ssf1.xml_;_
Set shell = CreateObject("Wscript.shell")
		shell.SendKeys "{PGDN}" @@ hightlight id_;_81_;_script infofile_;_ZIP::ssf19.xml_;_
		shell.SendKeys "{PGDN}"

JavaWindow("Sistema Integrador v1").JavaTab("Complete").JavaTab("Paso 1:").JavaTab("Formulario Virtual Nº").JavaEdit("401").Set 401
JavaWindow("Sistema Integrador v1").JavaTab("Complete").JavaTab("Paso 1:").JavaTab("Formulario Virtual Nº").JavaEdit("402").Set 402
JavaWindow("Sistema Integrador v1").JavaTab("Complete").JavaTab("Paso 1:").JavaTab("Formulario Virtual Nº").JavaEdit("403").Set 403

JavaWindow("Sistema Integrador v1").InsightObject("404").Click
	JavaWindow("Sistema Integrador v1").JavaButton("Agregar").Click
	JavaWindow("Sistema Integrador v1").InsightObject("InsightObject").Click
	JavaWindow("Sistema Integrador v1").InsightObject("Ruc").Click
	JavaWindow("Sistema Integrador v1").InsightObject("Numdoc").Type 10036940170
	wait 1
	JavaWindow("Sistema Integrador v1").InsightObject("Saldo").type 1234567
	wait 1
	JavaWindow("Sistema Integrador v1").JavaButton("Guardar").Click
	Set shell = CreateObject("Wscript.shell")
		shell.SendKeys "{ENTER}" @@ hightlight id_;_721335853_;_script infofile_;_ZIP::ssf25.xml_;_
	wait 2
	JavaWindow("Sistema Integrador v1").JavaButton("Cancelar").Click
	WAIT 2 @@ hightlight id_;_5_;_script infofile_;_ZIP::ssf27.xml_;_
	 
JavaWindow("Sistema Integrador v1").InsightObject("405").Click
	JavaWindow("Sistema Integrador v1").JavaButton("Agregar").Click
	JavaWindow("Sistema Integrador v1").InsightObject("InsightObject").Click
	JavaWindow("Sistema Integrador v1").InsightObject("Ruc").Click
	JavaWindow("Sistema Integrador v1").InsightObject("Numdoc").Type 10036940170
	wait 1
	JavaWindow("Sistema Integrador v1").InsightObject("Saldo").type 1234567
	wait 1
	JavaWindow("Sistema Integrador v1").JavaButton("Guardar").Click
	Set shell = CreateObject("Wscript.shell")
		shell.SendKeys "{ENTER}" @@ hightlight id_;_721335853_;_script infofile_;_ZIP::ssf25.xml_;_
	WAIT 2
	JavaWindow("Sistema Integrador v1").JavaButton("Cancelar").Click
	WAIT 2
	
JavaWindow("Sistema Integrador v1").InsightObject("407").Click
	JavaWindow("Sistema Integrador v1").JavaButton("Agregar").Click
	JavaWindow("Sistema Integrador v1").InsightObject("InsightObject").Click
	JavaWindow("Sistema Integrador v1").InsightObject("Ruc").Click
	JavaWindow("Sistema Integrador v1").InsightObject("Numdoc").Type 10036940170
	wait 1
	JavaWindow("Sistema Integrador v1").InsightObject("Saldo").type 1234567
	wait 1
	JavaWindow("Sistema Integrador v1").JavaButton("Guardar").Click
	Set shell = CreateObject("Wscript.shell")
		shell.SendKeys "{ENTER}" @@ hightlight id_;_721335853_;_script infofile_;_ZIP::ssf25.xml_;_
	WAIT 2
	JavaWindow("Sistema Integrador v1").JavaButton("Cancelar").Click
	WAIT 2
	
JavaWindow("Sistema Integrador v1").InsightObject("408").Click
	JavaWindow("Sistema Integrador v1").JavaButton("Agregar").Click
	JavaWindow("Sistema Integrador v1").InsightObject("InsightObject").Click
	JavaWindow("Sistema Integrador v1").InsightObject("Ruc").Click
	JavaWindow("Sistema Integrador v1").InsightObject("Numdoc").Type 10036940170
	wait 1
	JavaWindow("Sistema Integrador v1").InsightObject("Saldo").type 1234567
	wait 1
	JavaWindow("Sistema Integrador v1").JavaButton("Guardar").Click
	Set shell = CreateObject("Wscript.shell")
		shell.SendKeys "{ENTER}" @@ hightlight id_;_721335853_;_script infofile_;_ZIP::ssf25.xml_;_
	WAIT 2
	JavaWindow("Sistema Integrador v1").JavaButton("Cancelar").Click
	WAIT 2

JavaWindow("Sistema Integrador v1").JavaTab("Complete").JavaTab("Paso 1:").JavaTab("Formulario Virtual Nº").JavaEdit("406").Set 406
JavaWindow("Sistema Integrador v1").JavaTab("Complete").JavaTab("Paso 1:").JavaTab("Formulario Virtual Nº").JavaEdit("409").Set 409
JavaWindow("Sistema Integrador v1").JavaTab("Complete").JavaTab("Paso 1:").JavaTab("Formulario Virtual Nº").JavaEdit("410").Set 410
JavaWindow("Sistema Integrador v1").JavaTab("Complete").JavaTab("Paso 1:").JavaTab("Formulario Virtual Nº").JavaEdit("411").Set 411
JavaWindow("Sistema Integrador v1").JavaTab("Complete").JavaTab("Paso 1:").JavaTab("Formulario Virtual Nº").JavaEdit("414").Set 414
JavaWindow("Sistema Integrador v1").JavaTab("Complete").JavaTab("Paso 1:").JavaTab("Formulario Virtual Nº").JavaEdit("415").Set 415
JavaWindow("Sistema Integrador v1").JavaTab("Complete").JavaTab("Paso 1:").JavaTab("Formulario Virtual Nº").JavaEdit("416").Set 416
JavaWindow("Sistema Integrador v1").JavaTab("Complete").JavaTab("Paso 1:").JavaTab("Formulario Virtual Nº").JavaEdit("417").Set 417
JavaWindow("Sistema Integrador v1").JavaTab("Complete").JavaTab("Paso 1:").JavaTab("Formulario Virtual Nº").JavaEdit("418").Set 418
JavaWindow("Sistema Integrador v1").JavaTab("Complete").JavaTab("Paso 1:").JavaTab("Formulario Virtual Nº").JavaEdit("419").Set 419
JavaWindow("Sistema Integrador v1").JavaTab("Complete").JavaTab("Paso 1:").JavaTab("Formulario Virtual Nº").JavaEdit("420").Set 420
JavaWindow("Sistema Integrador v1").JavaTab("Complete").JavaTab("Paso 1:").JavaTab("Formulario Virtual Nº").JavaEdit("421").Set 421
JavaWindow("Sistema Integrador v1").JavaTab("Complete").JavaTab("Paso 1:").JavaTab("Formulario Virtual Nº").JavaEdit("422").Set 422
JavaWindow("Sistema Integrador v1").JavaTab("Complete").JavaTab("Paso 1:").JavaTab("Formulario Virtual Nº").JavaEdit("423").Set 423
JavaWindow("Sistema Integrador v1").JavaTab("Complete").JavaTab("Paso 1:").JavaTab("Formulario Virtual Nº").JavaEdit("424").Set 424

JavaWindow("Sistema Integrador v1").InsightObject("TOTAL").Click

	For Iterator = 1 To 15 Step 1
		Set shell = CreateObject("Wscript.Shell")
			shell.SendKeys "{PGUP}"
	Next @@ hightlight id_;_11_;_script infofile_;_ZIP::ssf22.xml_;_

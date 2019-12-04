JavaWindow("Sistema Integrador v1").JavaTab("Formulario Virtual N°").JavaTab("Sección Informativa").JavaTab("Información Complementaria").JavaButton("Agregar").Click
JavaWindow("Sistema Integrador v1").JavaList("Tipo de Socio :").Select "#3"
JavaWindow("Sistema Integrador v1").JavaList("Tipodoc").Select "#2"
JavaWindow("Sistema Integrador v1").JavaEdit("Numero de Documento :").Set 44394190
wait 3
'JavaWindow("Sistema Integrador v1").JavaEdit("Nombre o Razón Social").Set Razon
JavaWindow("Sistema Integrador v1").InsightObject("Calendario").Click
wait 2
Set shell = CreateObject("Wscript.shell")

	shell.SendKeys "{ENTER}"
wait 1
JavaWindow("Sistema Integrador v1").JavaList("País de residencia :").Select "#2"
'JavaWindow("Sistema Integrador v1").JavaEdit("Porcentaje de participación").Set 23
JavaWindow("Sistema Integrador v1").JavaEdit("Alquileres Pagados").Set 12
JavaWindow("Sistema Integrador v1").InsightObject("Calendario2").Click
wait 1
Set shell = CreateObject("Wscript.shell")
	shell.SendKeys "{ENTER}"
wait 1

JavaWindow("Sistema Integrador v1").JavaButton("Grabar").Click


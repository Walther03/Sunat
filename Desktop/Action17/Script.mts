JavaWindow("Sistema Integrador v1").InsightObject("Mas").Click @@ hightlight id_;_1_;_script infofile_;_ZIP::ssf2.xml_;_
JavaWindow("Sistema Integrador v1").InsightObject("PestañaVehiculos").Click @@ hightlight id_;_7_;_script infofile_;_ZIP::ssf3.xml_;_
gastos = "Si"
If gastos = "Si" Then
	JavaWindow("Sistema Integrador v1").InsightObject("Si").Click
	JavaWindow("Sistema Integrador v1").InsightObject("IngresoNetoAnual").Click
	JavaWindow("Sistema Integrador v1").InsightObject("Determinacion_Fecha").Click
	Set shell = CreateObject("Wscript.Shell")
				shell.SendKeys "{ENTER}"
	JavaWindow("Sistema Integrador v1").InsightObject("Determinacion_705").Click
	JavaWindow("Sistema Integrador v1").JavaEdit("Determinacion_705_707").Set 12
	JavaWindow("Sistema Integrador v1").JavaEdit("Determinacion_705_708").Set 12
	JavaWindow("Sistema Integrador v1").JavaEdit("Determinacion_705_709").Set 12
	JavaWindow("Sistema Integrador v1").JavaButton("Aceptar").Click
	
	JavaWindow("Sistema Integrador v1").InsightObject("Determinacion_706").Click
	JavaWindow("Sistema Integrador v1").InsightObject("Determinacion_706_2018").Click
	JavaWindow("Sistema Integrador v1").JavaEdit("ENERO").Set 1
	JavaWindow("Sistema Integrador v1").JavaEdit("FEBRERO").Set 2
	JavaWindow("Sistema Integrador v1").JavaEdit("MARZO").Set 3
	JavaWindow("Sistema Integrador v1").JavaEdit("ABRIL").Set 4
	JavaWindow("Sistema Integrador v1").JavaEdit("MAYO").Set 5
	JavaWindow("Sistema Integrador v1").JavaEdit("JUNIO").Set 6
	JavaWindow("Sistema Integrador v1").JavaEdit("JULIO").Set 7
	JavaWindow("Sistema Integrador v1").JavaEdit("AGOSTO").Set 8
	JavaWindow("Sistema Integrador v1").JavaEdit("SETIEMBRE").Set 9
	JavaWindow("Sistema Integrador v1").JavaEdit("OCTUBRE").Set 10
	JavaWindow("Sistema Integrador v1").JavaEdit("NOVIEMBRE").Set 11
	JavaWindow("Sistema Integrador v1").JavaEdit("DICIEMBRE").Set 12
	JavaWindow("Sistema Integrador v1").JavaButton("Aceptar").Click
	
	JavaWindow("Sistema Integrador v1").InsightObject("Determinacion_706_2019").Click
	JavaWindow("Sistema Integrador v1").JavaEdit("ENERO").Set 1
	JavaWindow("Sistema Integrador v1").JavaEdit("FEBRERO").Set 2
	JavaWindow("Sistema Integrador v1").JavaEdit("MARZO_2").Set 3
	JavaWindow("Sistema Integrador v1").JavaEdit("ABRIL_2").Set 4
	JavaWindow("Sistema Integrador v1").JavaEdit("MAYO").Set 5
	JavaWindow("Sistema Integrador v1").JavaEdit("JUNIO").Set 6
	JavaWindow("Sistema Integrador v1").JavaEdit("JULIO_2").Set 7
	JavaWindow("Sistema Integrador v1").JavaEdit("AGOSTO").Set 8
	JavaWindow("Sistema Integrador v1").JavaEdit("SETIEMBRE").Set 9
	JavaWindow("Sistema Integrador v1").JavaEdit("OCTUBRE").Set 10
	JavaWindow("Sistema Integrador v1").JavaEdit("NOVIEMBRE").Set 11
	JavaWindow("Sistema Integrador v1").JavaEdit("DICIEMBRE").Set 12
	JavaWindow("Sistema Integrador v1").JavaEdit("Promedio de ingresos").Set 12

	JavaWindow("Sistema Integrador v1").JavaButton("Aceptar").Click
	
	JavaWindow("Sistema Integrador v1").JavaButton("Aceptar").Click
	
	JavaWindow("Sistema Integrador v1").JavaButton("Aceptar").Click
	
	
	JavaWindow("Sistema Integrador v1").JavaTab("Regalías a Beneficiarios").JavaTab("Regalías a Beneficiarios").JavaTab("Regalías a Beneficiarios").JavaEdit("701").Set 1
	JavaWindow("Sistema Integrador v1").JavaTab("Regalías a Beneficiarios").JavaTab("Regalías a Beneficiarios").JavaTab("Regalías a Beneficiarios").JavaEdit("702").Set 2
	JavaWindow("Sistema Integrador v1").JavaTab("Regalías a Beneficiarios").JavaTab("Regalías a Beneficiarios").JavaTab("Regalías a Beneficiarios").JavaEdit("703").Set 3
	JavaWindow("Sistema Integrador v1").JavaTab("Regalías a Beneficiarios").JavaTab("Regalías a Beneficiarios").JavaTab("Regalías a Beneficiarios").JavaEdit("704").Set 4

	
End If



'JavaWindow("Sistema Integrador v1").InsightObject("IngresoNetoAnual").InsightObject("706").Click


	



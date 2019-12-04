

JavaWindow("Sistema Integrador v1").InsightObject("Pestaña_Alquiler").Click

Function Fun_Alquileres(alquiler)
	
	If alquiler = "Si" Then
		
		JavaWindow("Sistema Integrador v1").InsightObject("BtnSi").Click
		JavaWindow("Sistema Integrador v1").JavaTab("Paso 1:").JavaTab("Sección Informativa").JavaTab("Información Complementaria").JavaButton("Agregar").Click
		JavaWindow("Sistema Integrador v1").JavaList("Tipo de documento").Select "#3"
		JavaWindow("Sistema Integrador v1").JavaEdit("Numdoc").Set 44373946
		JavaWindow("Sistema Integrador v1").JavaEdit("Monto anual del alquiler").Set 349
		bien = "BIEN INMUEBLE DISTINTO"
		Select Case bien
	
			Case "BIEN INMUEBLE"
				JavaWindow("Sistema Integrador v1").InsightObject("Rdb_BienMueble").Click
				JavaWindow("Sistema Integrador v1").JavaList("Seleccione el bien").Select "#2"
				JavaWindow("Sistema Integrador v1").JavaEdit("Matricula").Set "AAAAAAAAAA"
		
			Case "BIEN INMUEBLE (PREDIOS)"
				JavaWindow("Sistema Integrador v1").InsightObject("Rdb_BienIPredio").Click
		
			Case "BIEN INMUEBLE DISTINTO"
				JavaWindow("Sistema Integrador v1").InsightObject("Rdb_InmuebleDistinto").Click
				JavaWindow("Sistema Integrador v1").JavaList("Seleccione el bien").Select "#1"
				JavaWindow("Sistema Integrador v1").JavaEdit("Partida Registral").Set "aaaaaa"

		End Select

		JavaWindow("Sistema Integrador v1").JavaButton("Aceptar").Click
		
	Else 
		
		if JavaWindow("Sistema Integrador v1").InsightObject("Alquile_Rdb_No").Exist = false then 

			JavaWindow("Sistema Integrador v1").InsightObject("No_1").Click
		
		Else 
			JavaWindow("Sistema Integrador v1").InsightObject("Alquile_Rdb_No").Click
		End If
	
	End If	
		
	
End Function


		







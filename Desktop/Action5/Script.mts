JavaWindow("Sistema Integrador v1").InsightObject("Pestaña_Amazonia").Click

AcogimientoAmazonica = "Si"
RegimenDomicilio	 = "Si"
TipoZona			 = "Zona 1"
ofi_registral		 = 1
tomo				 = 2
folio				 = 3
Asientos			 = 4
montofijo			 = 5

Call Fun_AcogimientoAmazonica(AcogimientoAmazonica)
Call Fun_MostrarMenu()


Function Fun_AcogimientoAmazonica(AcogimientoAmazonica)
		
	If AcogimientoAmazonica = "Si" Then
		
		JavaWindow("Sistema Integrador v1").JavaTab("Complete").JavaTab("Sección Informativa").JavaTab("Identificación").JavaRadioButton("Acogimiento_Si").Set
		Call Fun_RegimenDomicilio(RegimenDomicilio,TipoZona)
		
	Else 
	
		JavaWindow("Sistema Integrador v1").JavaTab("Complete").JavaTab("Sección Informativa").JavaTab("Identificación").JavaRadioButton("Acogimiento_No").Set
		Call Fun_Contribuyentes(Contribuyentes)
		
	End If
	
End Function
Function Fun_RegimenDomicilio(RegimenDomicilio,TipoZona)
	
	If RegimenDomicilio = "Si" Then
		
		JavaWindow("Sistema Integrador v1").JavaTab("Complete").JavaTab("Sección Informativa").JavaTab("Identificación").JavaRadioButton("RegimenDomicilio_Si").Set
	
		Select Case TipoZona
			Case "Zona 1"
				JavaWindow("Sistema Integrador v1").JavaTab("Complete").JavaTab("Sección Informativa").JavaTab("Identificación").JavaRadioButton("RegimenDomicilio_Zona 1").Set

			
			Case "Zona 2"
				JavaWindow("Sistema Integrador v1").JavaTab("Complete").JavaTab("Sección Informativa").JavaTab("Identificación").JavaRadioButton("RegimenDomicilio_Zona 2").Set

		End Select
		
		For Iterator = 1 To 15 Step 1
			Set shell = CreateObject("Wscript.Shell")
				shell.SendKeys "{PGDN}"
		Next
		JavaWindow("Sistema Integrador v1").InsightObject("RegimenDomicilio_Ubigeo").Click


		JavaWindow("Sistema Integrador v1").JavaList("RegimenDomicilio_Departamento:").Select "#1"
		JavaWindow("Sistema Integrador v1").JavaList("RegimenDomicilio_Provincia:").Select "#1"
		JavaWindow("Sistema Integrador v1").JavaList("RegimenDomicilio_Distrito:").Select "#1"
		JavaWindow("Sistema Integrador v1").JavaButton("RegimenDomicilio_Guardar").Click
		
		Call Fun_RegistrosPublicos(ofi_registral,tomo,folio,Asientos,montofijo)
		JavaWindow("Sistema Integrador v1").JavaTab("Complete").JavaTab("Sección Informativa").JavaTab("Identificación").JavaRadioButton("RegimenDomicilio_Si").Set
	Else
	
		JavaWindow("Sistema Integrador v1").JavaTab("Complete").JavaTab("Sección Informativa").JavaTab("Identificación").JavaRadioButton("RegimenDomicilio_No").Set

	End If
	
	
	
End Function
Function Fun_Contribuyentes(Contribuyentes)
	If Contribuyentes = "Si" Then
		JavaWindow("Sistema Integrador v1").JavaTab("Complete").JavaTab("Sección Informativa").JavaTab("Identificación").JavaRadioButton("ContribuyeAmazonia_Si").Set
		Call Fun_RegimenDomicilio(RegimenDomicilio,TipoZona)
	Else
		
	JavaWindow("Sistema Integrador v1").JavaTab("Complete").JavaTab("Sección Informativa").JavaTab("Identificación").JavaRadioButton("ContribuyeAmazonia_No").Set

	End If
	
End Function
Function Fun_RegistrosPublicos(ofi_registral,tomo,folio,Asientos,montofijo)
	
JavaWindow("Sistema Integrador v1").JavaTab("Complete").JavaTab("Sección Informativa").JavaTab("Identificación").JavaEdit("Reg_280").Set ofi_registral
JavaWindow("Sistema Integrador v1").JavaTab("Complete").JavaTab("Sección Informativa").JavaTab("Identificación").JavaEdit("Reg_281").Set tomo
JavaWindow("Sistema Integrador v1").JavaTab("Complete").JavaTab("Sección Informativa").JavaTab("Identificación").JavaEdit("Reg_282").Set folio
JavaWindow("Sistema Integrador v1").JavaTab("Complete").JavaTab("Sección Informativa").JavaTab("Identificación").JavaEdit("Reg_283").Set Asientos

JavaWindow("Sistema Integrador v1").JavaTab("Complete").JavaTab("Sección Informativa").JavaTab("Identificación").JavaEdit("Reg_223").Set montofijo

End Function
Function Fun_MostrarMenu 
	For Iterator = 1 To 15 Step 1
		Set shell = CreateObject("Wscript.Shell")
			shell.SendKeys "{PGUP}"
	Next
End Function
		
		

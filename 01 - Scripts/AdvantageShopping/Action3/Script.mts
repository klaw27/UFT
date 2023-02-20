'PRUEBA GIT

strIdentificador = DataTable("nombrePrueba", dtGlobalSheet)

If DataTable("ejecutarPrueba", dtGlobalSheet) = "ON" Then

	Parameter("Resultado") = AbrirNavegador()

	If Parameter("Resultado") Then
		Reporter.ReportEvent micDone, "NavegadorAdvantage", "El navegador se ha abierto correctamente"
		RunAction "CrearUsuario", oneIteration, Parameter("Resultado")
	End If
	
	If Parameter("Resultado") Then
		RunAction "RealizarCompra", oneIteration, Parameter("Resultado")
	End If
	
	If Parameter("Resultado") Then
		RunAction "BorrarUsuario", oneIteration, Parameter("Resultado")
	End If
	
	CerrarNavegador(Browser("Advantage Browser"))
	
End If

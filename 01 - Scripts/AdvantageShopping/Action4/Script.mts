

Parameter("Resultado") = VolverInicio()

If Parameter("Resultado") Then
	Parameter("Resultado") = RealizarCompra()
	If Parameter("Resultado") Then
		Reporter.ReportEvent micPass, "RealizarCompra", "La compra se ha realizado correctamente"
	Else
		Reporter.ReportEvent micFail, "RealizarCompra", "Ha ocurrido un error al realizar la compra"
	End If
End If



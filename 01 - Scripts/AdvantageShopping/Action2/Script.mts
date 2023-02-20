

Parameter("Resultado") = EliminarCuenta

If Parameter("Resultado") Then
	Reporter.ReportEvent micPass, "Funcion_EliminarCuenta", "La función eliminar cuenta ha finalizado correctamente"
Else
	Reporter.ReportEvent micFail, "Funcion_EliminarCuenta", "Ha ocurrido un error al eliminar el usuario"
	CapturaError Browser("Advantage Browser"), "Funcion_EliminarCuenta"
End If

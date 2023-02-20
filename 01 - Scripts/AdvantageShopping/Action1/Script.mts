

Parameter("Resultado") = CrearUsuario

If Parameter("Resultado") Then
	Reporter.ReportEvent micPass, "Funcion_CrearUsuario", "La función crear usuario ha finalizado correctamente"
Else
	Reporter.ReportEvent micFail, "Funcion_CrearUsuario", "Ha ocurrido un error al crear el usuario"
	CapturaError Browser("Advantage Browser"), "Funcion_CrearUsuario"
End If


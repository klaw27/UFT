'PRUEBA GIT - Pull cambios en la nube

'strIdentificador = DataTable("nombrePrueba", dtGlobalSheet)
'
'If DataTable("ejecutarPrueba", dtGlobalSheet) = "ON" Then
'
'	Parameter("Resultado") = AbrirNavegador()
'
'	If Parameter("Resultado") Then
'		Reporter.ReportEvent micDone, "NavegadorAdvantage", "El navegador se ha abierto correctamente"
'		RunAction "CrearUsuario", oneIteration, Parameter("Resultado")
'	End If
'	
'	If Parameter("Resultado") Then
'		RunAction "RealizarCompra", oneIteration, Parameter("Resultado")
'	End If
'	
'	If Parameter("Resultado") Then
'		RunAction "BorrarUsuario", oneIteration, Parameter("Resultado")
'	End If
'	
'	CerrarNavegador(Browser("Advantage Browser"))
'	
'End If



Set conexionBD = CreateObject("ADODB.Connection")
conexionBD.Open "provider=sqloledb; server=DESKTOP-8ELHMH0\SQLEXPRESS; user id=sa; password=prueba123; Database=dbUFT; Trusted_Connection=Yes;"


Set recordset = CreateObject("ADODB.Recordset")
query1 = "Select * from tablaPrueba"

recordset.Open query1, conexionBD
valor = recordset.Fields.Item(0)
msgbox (valor)

recordset.Close
conexionBD.Close

Set conexionBD = Nothing
Set recordset = Nothing

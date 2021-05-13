'******************************************************************************************
'** procesa exportacion de ventas del sistema rece                                        *
'** convierte a excel									                                                    *
'** http://msdn.microsoft.com/en-us/library/windows/desktop/ms709353%28v=vs.85%29.aspx    *
'**                                                                                       *
'******************************************************************************************

Option Explicit

'****** NO EDITAR POR DEBAJO DE ESTA LINEA  ****************************************
Dim objArgs
'indica que abriremos un fixed-width text file
Const xlFixedWidth = 2
'constante que indica el formato de las columnas a importar
' ver xlColumnDataType  en http://msdn.microsoft.com/en-us/library/aa221100(office.11).aspx
Const xlGeneralFormat = 1
Const xlTextFormat = 2
Const xlYMDFormat  = 5
Const xlDMYFormat = 4
Const xlYDMFormat=8
Const xlSkipColumn = 9
Const xlToRight = -4161
Const frmtCurrency = "$#,##0.00"

'Columna que tiene el importe total
Const totalRow = 9
'Columna que tiene el tipo
Const typeRow = 2

'constante que indica cual es la primer fila a importar
Const fstline = 1
'Nombre del archivo de rendicion y archivo de salida
Dim infilename,xlfilename
Dim xlHeader,xlHField,indx, indj
'Variables Execel
Dim oExcel, oWb
Dim oRange
Dim oSheet, oSheetTipos


xlHeader =Array("Fecha de comprobante", "Tipo de comprobante", "Tipo TxT", "Punto de venta", "Numero de comprobante",_
"Numero de comprobante hasta", "Codigo de documento del comprador", "Numero de identificacion del comprador",_
"Apellido y nombre del comprador",_
"Importe total de la operacion",_
"Importe en Pesos",_
"Importe total de conceptos que no integran el precio neto gravado", "Percepcion a no categorizados",_
"Importe operaciones exentas", "Importe de percepciones o pagos a cuenta de impuestos nacionales",_
"Importe de percepciones de ingresos brutos", "Importe de percepciones impuestos municipales",_
"Importe impuestos internos", "Codigo de Moneda", "Tipo de cambio", "Cantidad de alicuotas de IVA",_
"Codigo de operacion", "Otros Tributos", "Fecha de vencimiento de pago")

'Procesamiento Parameters
Set objArgs = WScript.Arguments

IF (objArgs.Count)<1 then
	Wscript.echo "debe ingresar el archivo de ventas como parametro"
	Wscript.quit
end if
'Archivo de rendicion enviado como parametro
infilename = objArgs(0)


'Generacion de los nombres de los archivos
Dim objFSO, objFile
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.GetFile(infilename)

'Nombre del archivo de salida
xlfilename = objFSO.GetParentFolderName(objFile) & "\" & objFSO.GetBaseName(objFile) & ".xlsx"

Set objFSO = nothing
Set objFile = nothing

Set oExcel = CreateObject("Excel.Application")
oExcel.Visible = True

oExcel.Workbooks.OpenText infilename,,fstline,xlFixedWidth,,,,,,,,, _
  Array(_
Array(0, xlYMDFormat),_
Array(8, xlGeneralFormat),_
Array(11, xlGeneralFormat),_
Array(16, xlGeneralFormat),_
Array(36, xlGeneralFormat),_
Array(56, xlGeneralFormat),_
Array(58, xlTextFormat),_
Array(78, xlTextFormat),_
Array(108, xlGeneralFormat),_
Array(123, xlGeneralFormat),_
Array(138, xlGeneralFormat),_
Array(153, xlGeneralFormat),_
Array(168, xlGeneralFormat),_
Array(183, xlGeneralFormat),_
Array(198, xlGeneralFormat),_
Array(213, xlGeneralFormat),_
Array(228, xlTextFormat),_
Array(231, xlGeneralFormat),_
Array(241, xlGeneralFormat),_
Array(242, xlTextFormat),_
Array(243, xlGeneralFormat),_
Array(258, xlYMDFormat))

Set oSheet = oExcel.ActiveSheet
Set oWb = oExcel.ActiveWorkbook
Set oSheetTipos = oWb.Worksheets.Add(, oWb.Worksheets(oWb.Worksheets.Count))

Dim tiposData
tiposData= Array(Array("Codigo","Descripcion","DC"),_
Array("001","FACTURAS A",1),_
Array("002","NOTAS DE DEBITO A",1),_
Array("003","NOTAS DE CREDITO A",-1),_
Array("004","RECIBOS A",0),_
Array("005","NOTAS DE VENTA AL CONTADO A",0),_
Array("006","FACTURAS B",1),_
Array("007","NOTAS DE DEBITO B",1),_
Array("008","NOTAS DE CREDITO B",-1),_
Array("009","RECIBOS B",0),_
Array("010","NOTAS DE VENTA AL CONTADO B",0),_
Array("011","FACTURAS C",1),_
Array("012","NOTAS DE DEBITO C",1),_
Array("013","NOTAS DE CREDITO C",-1),_
Array("015","RECIBOS C",0),_
Array("016","NOTAS DE VENTA AL CONTADO C",0),_
Array("017","LIQUIDACION DE SERVICIOS PUBLICOS CLASE A",0),_
Array("018","LIQUIDACION DE SERVICIOS PUBLICOS CLASE B",0),_
Array("019","FACTURAS DE EXPORTACION",1),_
Array("020","NOTAS DE DEBITO POR OPERACIONES CON EL EXTERIOR",1),_
Array("021","NOTAS DE CREDITO POR OPERACIONES CON EL EXTERIOR",-1),_
Array("022","FACTURAS - PERMISO EXPORTACION SIMPLIFICADO - DTO. 855/97",0),_
Array("023","COMPROBANTES “A” DE COMPRA PRIMARIA PARA EL SECTOR PESQUERO MARITIMO",0),_
Array("024","COMPROBANTES “A” DE CONSIGNACION PRIMARIA PARA EL SECTOR PESQUERO MARITIMO",0),_
Array("025","COMPROBANTES “B” DE COMPRA PRIMARIA PARA EL SECTOR PESQUERO MARITIMO",0),_
Array("026","COMPROBANTES “B” DE CONSIGNACION PRIMARIA PARA EL SECTOR PESQUERO MARITIMO",0),_
Array("027","LIQUIDACION UNICA COMERCIAL IMPOSITIVA CLASE A",0),_
Array("028","LIQUIDACION UNICA COMERCIAL IMPOSITIVA CLASE B",0),_
Array("029","LIQUIDACION UNICA COMERCIAL IMPOSITIVA CLASE C",0),_
Array("030","COMPROBANTES DE COMPRA DE BIENES USADOS",0),_
Array("031","MANDATO - CONSIGNACION",0),_
Array("032","COMPROBANTES PARA RECICLAR MATERIALES",0),_
Array("033","LIQUIDACION PRIMARIA DE GRANOS",0),_
Array("034","COMPROBANTES A DEL APARTADO A  INCISO F)  R.G. N°  1415",0),_
Array("035","COMPROBANTES B DEL ANEXO I, APARTADO A, INC. F), R.G. N° 1415",0),_
Array("036","COMPROBANTES C DEL Anexo I, Apartado A, INC.F), R.G. N° 1415",0),_
Array("037","NOTAS DE DEBITO O DOCUMENTO EQUIVALENTE QUE CUMPLAN CON LA R.G. N° 1415",0),_
Array("038","NOTAS DE CREDITO O DOCUMENTO EQUIVALENTE QUE CUMPLAN CON LA R.G. N° 1415",0),_
Array("039","OTROS COMPROBANTES A QUE CUMPLEN CON LA R G  1415",0),_
Array("040","OTROS COMPROBANTES B QUE CUMPLAN CON LA R.G. N° 1415",0),_
Array("041","OTROS COMPROBANTES C QUE CUMPLAN CON LA R.G. N° 1415",0),_
Array("043","NOTA DE CREDITO LIQUIDACION UNICA COMERCIAL IMPOSITIVA CLASE B",0),_
Array("044","NOTA DE CREDITO LIQUIDACION UNICA COMERCIAL IMPOSITIVA CLASE C",0),_
Array("045","NOTA DE DEBITO LIQUIDACION UNICA COMERCIAL IMPOSITIVA CLASE A",0),_
Array("046","NOTA DE DEBITO LIQUIDACION UNICA COMERCIAL IMPOSITIVA CLASE B",0),_
Array("047","NOTA DE DEBITO LIQUIDACION UNICA COMERCIAL IMPOSITIVA CLASE C",0),_
Array("048","NOTA DE CREDITO LIQUIDACION UNICA COMERCIAL IMPOSITIVA CLASE A",0),_
Array("049","COMPROBANTES DE COMPRA DE BIENES NO REGISTRABLES A CONSUMIDORES FINALES",0),_
Array("050","RECIBO FACTURA A  REGIMEN DE FACTURA DE CREDITO ",0),_
Array("051","FACTURAS M",1),_
Array("052","NOTAS DE DEBITO M",1),_
Array("053","NOTAS DE CREDITO M",-1),_
Array("054","RECIBOS M",0),_
Array("055","NOTAS DE VENTA AL CONTADO M",0),_
Array("056","COMPROBANTES M DEL ANEXO I  APARTADO A  INC F) R.G. N° 1415",0),_
Array("057","OTROS COMPROBANTES M QUE CUMPLAN CON LA R.G. N° 1415",0),_
Array("058","CUENTAS DE VENTA Y LIQUIDO PRODUCTO M",0),_
Array("059","LIQUIDACIONES M",0),_
Array("060","CUENTAS DE VENTA Y LIQUIDO PRODUCTO A",0),_
Array("061","CUENTAS DE VENTA Y LIQUIDO PRODUCTO B",0),_
Array("063","LIQUIDACIONES A",0),_
Array("064","LIQUIDACIONES B",0),_
Array("066","DESPACHO DE IMPORTACION",0),_
Array("068","LIQUIDACION C",0),_
Array("070","RECIBOS FACTURA DE CREDITO",0),_
Array("080","INFORME DIARIO DE CIERRE (ZETA) - CONTROLADORES FISCALES",0),_
Array("081","TIQUE FACTURA A   ",0),_
Array("082","TIQUE FACTURA B",0),_
Array("083","TIQUE",0),_
Array("088","REMITO ELECTRONICO",0),_
Array("089","RESUMEN DE DATOS",0),_
Array("090","OTROS COMPROBANTES - DOCUMENTOS EXCEPTUADOS - NOTAS DE CREDITO",0),_
Array("091","REMITOS R",0),_
Array("099","OTROS COMPROBANTES QUE NO CUMPLEN O ESTÁN EXCEPTUADOS DE LA R.G. 1415 Y SUS MODIF ",0),_
Array("110","TIQUE NOTA DE CREDITO ",0),_
Array("111","TIQUE FACTURA C",0),_
Array("112"," TIQUE NOTA DE CREDITO A",0),_
Array("113","TIQUE NOTA DE CREDITO B",0),_
Array("114","TIQUE NOTA DE CREDITO C",0),_
Array("115","TIQUE NOTA DE DEBITO A",0),_
Array("116","TIQUE NOTA DE DEBITO B",0),_
Array("117","TIQUE NOTA DE DEBITO C",0),_
Array("118","TIQUE FACTURA M",0),_
Array("119","TIQUE NOTA DE CREDITO M",0),_
Array("120","TIQUE NOTA DE DEBITO M",0),_
Array("201","FACTURA DE CRÉDITO ELECTRÓNICA MiPyMEs (FCE) A",0),_
Array("202","NOTA DE DEBITO ELECTRÓNICA MiPyMEs (FCE) A",0),_
Array("203","NOTA DE CREDITO ELECTRÓNICA MiPyMEs (FCE) A",0),_
Array("206","FACTURA DE CRÉDITO ELECTRÓNICA MiPyMEs (FCE) B",0),_
Array("207","NOTA DE DEBITO ELECTRÓNICA MiPyMEs (FCE) B",0),_
Array("208","NOTA DE CREDITO ELECTRÓNICA MiPyMEs (FCE) B",0),_
Array("211","FACTURA DE CRÉDITO ELECTRÓNICA MiPyMEs (FCE) C",0),_
Array("212","NOTA DE DEBITO ELECTRÓNICA MiPyMEs (FCE) C",0),_
Array("213","NOTA DE CREDITO ELECTRÓNICA MiPyMEs (FCE) C",0),_
Array("331","LIQUIDACION SECUNDARIA DE GRANOS",0),_
Array("332","CERTIFICACION ELECTRONICA (GRANOS)",0),_
Array("995","REMITO ELECTRÓNICO CÁRNICO ",0))
With oSheetTipos
	.Name = "Tipos"
	For indx = 0 to UBound(tiposData)
		For indj = 0 to UBound(tiposData(indx))
			.Cells(indx+1, indj+1) = tiposData(indx)(indj)
		Next
	Next
	.Range(.Cells(1,1), .Cells(UBound(tiposData)+1,  UBound(tiposData(0))+1  )).Name= "tipos"
End With
oSheet.Activate



'Borro el footer
' Set oRange = oSheet.Range("A" & oSheet.UsedRange.Rows.Count).EntireRow
' oRange.Delete

'Agrego importe
Dim strFormulas,oFstAvailable, oFstFinal, oLastAvailable
With oSheet
	' Agrego columna de tipo
	.Columns(totalRow+1).Insert xlToRight
	.Cells(1,totalRow).EntireColumn.Name="Tipo"

	'Agrego la columna para el calculo
	.Columns(totalRow+1).Insert xlToRight

	strFormulas = Array ("=(Importe_original/100) * VLOOKUP(B1,tipos,3,0)")
	'Celda donde inserto
	Set oFstAvailable = .Cells(1, totalRow+1)
	'Defino el nombre de la columna importe
	.Cells(1,totalRow).EntireColumn.Name="Importe_original"

	'Asigno formulas al rango
	.Range(oFstAvailable,oFstAvailable).Formula = strFormulas
	Set oLastAvailable = .Cells(.UsedRange.Rows.Count,totalRow+1)

	' Agrego columna de tipo
	' ================================================
	.Columns(typeRow+1).Insert xlToRight
	.Cells(1,typeRow).EntireColumn.Name="Tipo"
	.Range(.Cells(1, typeRow+1),.Cells(1, typeRow+1)).Formula = "=VLOOKUP(B1,tipos,2,0)"
	if .UsedRange.Rows.Count > 1 Then
		.Range(.Cells(1, typeRow+1),.Cells(.UsedRange.Rows.Count, typeRow+1)).FillDown
		.Range(oFstAvailable,oLastAvailable).FillDown
	end if
	' ================================================

	' Formato
	.Range(oFstAvailable,oLastAvailable).NumberFormat = frmtCurrency
	' Total Final
	Set oFstFinal = .Cells(.UsedRange.Rows.Count+1,totalRow+1)
	.Range(oFstAvailable,oLastAvailable).Name= "importe_a_sumar"
	.Range(oFstFinal, oFstFinal).Formula = "=SUM(importe_a_sumar)"
	.Range(oFstFinal, oFstFinal).NumberFormat = frmtCurrency


End With

'Inserto el header
Set oRange = oSheet.Range("A1").EntireRow
oRange.Insert
For indx=0 to UBound(xlHeader)
	oSheet.Cells(1, indx+1).Value = xlHeader(indx)
Next

'Muestro y grabo el excel
oExcel.ActiveWorkbook.SaveAs xlfilename

Set oRange=nothing
Set oSheet= nothing
Set oExcel=nothing



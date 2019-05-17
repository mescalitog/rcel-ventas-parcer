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

'constante que indica cual es la primer fila a importar
Const fstline = 1
'Nombre del archivo de rendicion y archivo de salida
Dim infilename,xlfilename
Dim xlHeader,xlHField,indx
'Variables Execel 
Dim oExcel
Dim oRange
Dim oSheet


xlHeader =Array("Fecha de comprobante", "Tipo de comprobante", "Punto de venta", "Numero de comprobante",_
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
'Borro el footer
' Set oRange = oSheet.Range("A" & oSheet.UsedRange.Rows.Count).EntireRow
' oRange.Delete

'Agrego importe
Dim strFormulas,oFstAvailable, oFstFinal, oLastAvailable
With oSheet
	'Agrego la columna para el calculo
	.Columns(totalRow+1).Insert xlToRight

	strFormulas = Array ("=Importe_original/100")

	'Celda donde inserto
	Set oFstAvailable = .Cells(1, totalRow+1)
	'Defino el nombre de la columna importe
	.Cells(1,totalRow).EntireColumn.Name="Importe_original"
	
	'Asigno formulas al rango
	.Range(oFstAvailable,oFstAvailable).Formula = strFormulas	
	Set oLastAvailable = .Cells(.UsedRange.Rows.Count,totalRow+1)
	'FillDown
	.Range(oFstAvailable,oLastAvailable).FillDown
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



using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Ajustador_de_formatos_Excel_de_movimientos_bancarios.BankFormatsFixerFunctions
{
    internal class Exterior
    {

        ExcelModifyFunctions modifyFunctions = new ExcelModifyFunctions();

        public int BankValidator(string rutaArchivo)
        {

            try
            {
                using (FileStream archivo = new FileStream(rutaArchivo, FileMode.Open))
                {
                    IWorkbook libro = new XSSFWorkbook(archivo);
                    string nombreHoja = libro.GetSheetAt(0).SheetName;
                    ISheet hoja = libro.GetSheet(nombreHoja);

                    IRow fila = hoja.GetRow(0);
                    if (fila != null)
                    {
                        string nCuenta = ExcelModifyFunctions.getValueCellString(fila.GetCell(1));
                        string validacion = ExcelModifyFunctions.getValueCellString(fila.GetCell(15));



                        if (validacion.Equals("Archivo modificado, Exterior"))
                        {
                            // CASO 2, EL DOCUMENTO SI ES UN FORMATO DE CONSULTA DE MOVIMIENTOS BCO, MODIFICADO

                            return 2;
                        }
                        else if (nCuenta.Contains("0115"))
                        {
                            // CASO 1, EL DOCUMENTO SI ES UN FORMATO DE CONSULTA DE MOVIMIENTOS BCO, SIN MODIFICAR
                            return 1;
                        }
                        else
                        {
                            // CASO 3, EL DOCUMENTO NO ES UN FORMATO DE CONSULTA DE MOVIMIENTOS BCO Exterior
                            return 0;
                        }

                    }


                }
            }
            catch (Exception ex)
            {

                MessageBox.Show("Error al verificar archivo: " + ex.Message, "ATENCIÓN", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return 0;
            }

            return 0;

        }

        public void fixFormat(TextBox ExcelFilePath)
        {

            List<string> listaColumnaFechas = modifyFunctions.CopyDateColumnsAsStrings(ExcelFilePath.Text, 0, 1);

            //Borrando las columnas con problemas de formato del banco (no cambian correctamente de General - Número)
            //y moviendo los datos

            modifyFunctions.MoveColumnsCaseBVnzlaBExterior(ExcelFilePath.Text, 4, 7, 4, 0);
            modifyFunctions.MoveColumnsCaseBVnzlaBExterior(ExcelFilePath.Text, 6, 9, 4, 0);
            modifyFunctions.MoveColumnsCaseBVnzlaBExterior(ExcelFilePath.Text, 7, 4, 4, 0);
            modifyFunctions.MoveColumnsCaseBVnzlaBExterior(ExcelFilePath.Text, 9, 6, 4, 0);

            //Limpiando columnas que ya no se utilizarán

            modifyFunctions.CleanColumn(ExcelFilePath.Text, 7, 0);
            modifyFunctions.CleanColumn(ExcelFilePath.Text, 9, 0);

            // Modificando las columnas para separar ingresos y egresos y borrar los símbolos + y - que vienen del banco

            MoveNumberRelatedtoSimbolCaseBancoExterior(ExcelFilePath.Text, 5, 4);

            //Ajustar Columnas para mejorar la estética
            modifyFunctions.AdjustColumnWidth(ExcelFilePath.Text, 1, 15, 0);
            modifyFunctions.AdjustColumnWidth(ExcelFilePath.Text, 2, 15, 0);
            modifyFunctions.AdjustColumnWidth(ExcelFilePath.Text, 3, 15, 0);
            modifyFunctions.AdjustColumnWidth(ExcelFilePath.Text, 4, 35, 0);
            modifyFunctions.AdjustColumnWidth(ExcelFilePath.Text, 5, 15, 0);
            modifyFunctions.AdjustColumnWidth(ExcelFilePath.Text, 6, 15, 0);
            modifyFunctions.AdjustColumnWidth(ExcelFilePath.Text, 7, 15, 0);

            //Moviendo antes de insertar
            //Columna total
            modifyFunctions.MoveColumnsCaseBVnzlaBExterior(ExcelFilePath.Text, 6, 7, 4, 0);
            //Columna egresos
            modifyFunctions.MoveColumnsCaseBVnzlaBExterior(ExcelFilePath.Text, 5, 6, 2, 0);
            //Columna ingresos
            modifyFunctions.MoveColumnsCaseBVnzlaBExterior(ExcelFilePath.Text, 4, 5, 2, 0);
            //Columna descripciones
            modifyFunctions.MoveColumnsCaseBVnzlaBExterior(ExcelFilePath.Text, 3, 4, 5, 0);
            //Columna referencias
            modifyFunctions.MoveColumnsCaseBVnzlaBExterior(ExcelFilePath.Text, 2, 3, 5, 0);

            //Dándole formato a las columnas numéricas, punto millar, coma decimal

            modifyFunctions.FormatNumericColumn(ExcelFilePath.Text, 5, 0);
            modifyFunctions.FormatNumericColumn(ExcelFilePath.Text, 6, 0);
            modifyFunctions.FormatNumericColumn(ExcelFilePath.Text, 7, 0);
            modifyFunctions.FormatNumericColumn(ExcelFilePath.Text, 8, 0);

            //Insertando columna de fecha de validación
            modifyFunctions.InsertColumnBetweenTwoVersionC3(ExcelFilePath.Text, 2, 0);

            //Cambio de columnas fecha y descripción
            modifyFunctions.MoveColumnsCaseBVnzlaBExterior(ExcelFilePath.Text, 1, 8, 5, 0);

            //Columna referencia
            modifyFunctions.MoveColumnsCaseBVnzlaBExterior(ExcelFilePath.Text, 4, 3, 5, 0);
            //Columna descripción
            modifyFunctions.MoveColumnsCaseBVnzlaBExterior(ExcelFilePath.Text, 8, 4, 5, 0);
            modifyFunctions.CleanColumn(ExcelFilePath.Text, 8, 0);

            //Reparando columna de referencia
            modifyFunctions.ConvertColumnToGeneral(ExcelFilePath.Text, 3, 0);

            //Trabajando la columna fecha (formato y posición)
            //Pegando columna fecha reparada
            modifyFunctions.changeCellTextFromListInTheSameOrder(ExcelFilePath.Text, 0, 0, listaColumnaFechas);


            //Ajustando formato de fecha
            modifyFunctions.ConvertColumnToGeneral(ExcelFilePath.Text, 1, 0);
            ChangeDateFormatCaseExteriorPrueba(ExcelFilePath.Text, 1, 0, 0);


            //Ajustando etiqueta de fecha de validación
            modifyFunctions.ChangeCellTextWithFormatAndStyle(ExcelFilePath.Text, 0, 1, "Fecha de validación");
            modifyFunctions.AdjustColumnWidth(ExcelFilePath.Text, 2, 20, 0);

            //Marcando la fecha de modificación para validar que el archivo ya fue manipulado
            modifyFunctions.ChangeCellTextWithFormatAndStyle(ExcelFilePath.Text, 0, 16, "Archivo modificado, Exterior");
            modifyFunctions.ChangeCellTextWithFormatAndStyle(ExcelFilePath.Text, 0, 17, "Fecha de modificación:" + DateTime.Now.ToString());

            //Reparando formato de las celdas en blanco (para que no se dañe la fórmula)
            modifyFunctions.replaceEmptyCellsWithZero(ExcelFilePath.Text, 0, 4);
            modifyFunctions.replaceEmptyCellsWithZero(ExcelFilePath.Text, 0, 5);
            modifyFunctions.replaceEmptyCellsWithZero(ExcelFilePath.Text, 0, 6);

            MessageBox.Show("Ajustes realizados exitosamente", "Proceso finalizado", MessageBoxButtons.OK, MessageBoxIcon.Information);



        }

        public void ChangeDateFormatCaseExteriorPrueba(string rutaArchivo, int columnaFecha, int nHoja, int filaIndexStart)
        {
            try
            {
                using (FileStream archivo = new FileStream(rutaArchivo, FileMode.Open))
                {
                    IWorkbook libro = new XSSFWorkbook(archivo);
                    string nombreHoja = libro.GetSheetAt(nHoja).SheetName;
                    ISheet hoja = libro.GetSheet(nombreHoja);

                    for (int filaIndex = filaIndexStart; filaIndex <= hoja.LastRowNum; filaIndex++)
                    {
                        IRow fila = hoja.GetRow(filaIndex);
                        if (fila != null)
                        {
                            ICell celdaFecha = fila.GetCell(columnaFecha - 1);
                            if (celdaFecha != null)
                            {
                                if (celdaFecha.CellType == CellType.Numeric && DateUtil.IsCellDateFormatted(celdaFecha))
                                {
                                    // Manejar celdas numéricas (fechas de Excel)
                                    double valorNumerico = celdaFecha.NumericCellValue;
                                    if (DateUtil.IsValidExcelDate(valorNumerico))
                                    {
                                        DateTime fechaBase = new DateTime(1899, 12, 30);
                                        DateTime fecha = fechaBase.AddDays(valorNumerico);
                                        string fechaString = fecha.ToString("dd/MM/yyyy");
                                        celdaFecha.SetCellValue(fechaString);
                                        celdaFecha.SetCellType(CellType.String); // Cambiar a tipo String
                                    }
                                }
                                else
                                {
                                    // Manejar celdas de cadena (dd/M/yyyy)
                                    celdaFecha.SetCellType(CellType.String);
                                    string valorFecha = modifyFunctions.getCellValueAsStringII(celdaFecha);
                                    if (!string.IsNullOrEmpty(valorFecha) && (valorFecha.Length == 8 || valorFecha.Length == 9 || valorFecha.Length == 10)) // Ajuste para manejar diferentes longitudes
                                    {
                                        try
                                        {
                                            DateTime fecha;
                                            if (DateTime.TryParseExact(valorFecha, "dd/M/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out fecha) ||
                                                DateTime.TryParseExact(valorFecha, "d/M/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out fecha) ||
                                                DateTime.TryParseExact(valorFecha, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out fecha)) // intento con el formato de salida
                                            {
                                                string fechaString = fecha.ToString("dd/MM/yyyy");
                                                celdaFecha.SetCellValue(fechaString);
                                            }
                                            else
                                            {
                                                Console.WriteLine($"Formato de fecha inválido en la fila {filaIndex + 1}: {valorFecha}");
                                            }
                                        }
                                        catch (FormatException ex)
                                        {
                                            Console.WriteLine($"Formato de fecha inválido en la fila {filaIndex + 1}: {ex.Message}");
                                        }
                                    }
                                }
                            }
                        }
                    }

                    using (FileStream archivoSalida = new FileStream(rutaArchivo, FileMode.Create))
                    {
                        libro.Write(archivoSalida);
                    }

                    Console.WriteLine("Formato de fecha cambiado exitosamente.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error al cambiar formato de fecha: {ex.Message}");
            }
        }

    /*
    * 
    * Esta función permite mover los números negativos de una columna, en función del signo contenido en su
    * correlativo en la columna justo a su derecha. Fue usada para ajustar el formato de banco exterior, y que 
    * las cantidades positivas y negativas no estén juntas en una columnas, siendo diferenciadas solo por esos
    * símbolos positivos y negativos en la columna a la derecha. 
    * 
    * */

        public void MoveNumberRelatedtoSimbolCaseBancoExterior(string rutaArchivo, int columnaSimbolos, int columnaNumeros)
        {
            try
            {
                using (FileStream archivo = new FileStream(rutaArchivo, FileMode.Open))
                {
                    IWorkbook libro = new XSSFWorkbook(archivo);
                    string nombreHoja = libro.GetSheetAt(0).SheetName;
                    ISheet hoja = libro.GetSheet(nombreHoja);

                    // Obtener el estilo de la celda origen (solo una vez)
                    ICell celdaOrigenEjemplo = hoja.GetRow(0).GetCell(1); // Celda de ejemplo para obtener el estilo
                    ICellStyle? estiloOrigen = null;

                    if (celdaOrigenEjemplo != null)
                    {
                        estiloOrigen = celdaOrigenEjemplo.CellStyle;
                        modifyFunctions.CopyCellStyle(celdaOrigenEjemplo.CellStyle, libro);
                    }

                    ICellStyle estiloCero = libro.CreateCellStyle();
                    estiloCero = estiloOrigen;

                    estiloCero.DataFormat = libro.CreateDataFormat().GetFormat("0;-0;;@");



                    for (int filaIndex = 0; filaIndex <= hoja.LastRowNum; filaIndex++)
                    {
                        IRow fila = hoja.GetRow(filaIndex);
                        if (fila != null)
                        {
                            ICell celdaSimbolo = fila.GetCell(columnaSimbolos - 1);
                            ICell celdaNumero = fila.GetCell(columnaNumeros - 1);

                            if (celdaSimbolo != null && celdaNumero != null)
                            {
                                string simbolo = celdaSimbolo.StringCellValue?.Trim();

                                if (simbolo == "+")
                                {
                                    // Dejar la celda en blanco
                                    celdaSimbolo.SetCellValue(0);
                                    celdaSimbolo.CellStyle = estiloCero;
                                }
                                else if (simbolo == "-")
                                {
                                    // Mover el número a la columna de símbolos
                                    if (celdaNumero.CellType == CellType.Numeric)
                                    {
                                        celdaSimbolo.SetCellValue(celdaNumero.NumericCellValue);
                                    }
                                    else if (celdaNumero.CellType == CellType.String && double.TryParse(celdaNumero.StringCellValue, out double valorNumerico))
                                    {
                                        celdaSimbolo.SetCellValue(valorNumerico);
                                    }

                                    // Limpiar la celda numérica
                                    celdaNumero.SetCellValue(0);
                                    celdaNumero.CellStyle = estiloCero;
                                }
                            }
                        }
                    }

                    // Guardar los cambios
                    using (FileStream archivoSalida = new FileStream(rutaArchivo, FileMode.Create))
                    {
                        libro.Write(archivoSalida);
                    }
                }
                Console.WriteLine($"Columnas procesadas exitosamente.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error al procesar columnas: {ex.Message}");
            }
        }



    }
}

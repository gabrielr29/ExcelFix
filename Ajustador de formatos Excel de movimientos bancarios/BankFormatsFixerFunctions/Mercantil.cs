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
    internal class Mercantil
    {
               
        public void ChangeDateFormatCaseMercantil(string rutaArchivo, int columnaFecha, int nHoja, int filaIndexStart)
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
                                // Forzar la lectura de la celda como cadena
                                celdaFecha.SetCellType(CellType.String);

                                string valorFecha = celdaFecha.StringCellValue;
                                if (!string.IsNullOrEmpty(valorFecha) && (valorFecha.Length == 7 || valorFecha.Length == 8))
                                {
                                    try
                                    {
                                        DateTime fecha;
                                        if (valorFecha.Length == 7)
                                        {
                                            fecha = DateTime.ParseExact("0" + valorFecha, "ddMMyyyy", CultureInfo.InvariantCulture);
                                        }
                                        else
                                        {
                                            fecha = DateTime.ParseExact(valorFecha, "ddMMyyyy", CultureInfo.InvariantCulture);
                                        }

                                        celdaFecha.SetCellValue(fecha);
                                        ICellStyle estiloFecha = libro.CreateCellStyle();
                                        IDataFormat formatoFecha = libro.CreateDataFormat();
                                        estiloFecha.DataFormat = formatoFecha.GetFormat("dd/MM/yyyy");
                                        celdaFecha.CellStyle = estiloFecha;
                                        celdaFecha.SetCellType(CellType.Numeric);
                                    }
                                    catch (FormatException ex)
                                    {
                                        Console.WriteLine($"Formato de fecha inválido en la fila {filaIndex + 1}: {ex.Message}");
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


        public int bankValidator(string rutaArchivo)
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

                        string validacion = ExcelModifyFunctions.getValueCellString(fila.GetCell(15));
                        string descripcion = ExcelModifyFunctions.getValueCellString(fila.GetCell(6)); // Normalizar descripción


                        // Crear una clave única para la fila


                        if (descripcion.Equals("NC") || descripcion.Equals("ND") || descripcion.Equals("DP") || descripcion.Equals("SF"))
                        {
                            // CASO 1, EL DOCUMENTO SI ES UN FORMATO DE CONSULTA DE MOVIMIENTOS BCO, SIN MODIFICAR
                            return 1;
                        }
                        else if (validacion.Equals("Archivo modificado, Mercantil"))
                        {
                            // CASO 2, EL DOCUMENTO SI ES UN FORMATO DE CONSULTA DE MOVIMIENTOS BCO, MODIFICADO

                            return 2;
                        }
                        else
                        {
                            // CASO 3, EL DOCUMENTO NO ES UN FORMATO DE CONSULTA DE MOVIMIENTOS BCO Mercantil
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

        public int bankValidatorNewVersion(string rutaArchivo)
        {

            try
            {
                using (FileStream archivo = new FileStream(rutaArchivo, FileMode.Open))
                {
                    IWorkbook libro = new XSSFWorkbook(archivo);
                    string nombreHoja = libro.GetSheetAt(0).SheetName;
                    ISheet hoja = libro.GetSheet(nombreHoja);

                    IRow fila = hoja.GetRow(3);
                    if (fila != null)
                    {

                       string cellTittleNT = ExcelModifyFunctions.getValueCellString(fila.GetCell(2)); // Normalizar descripción
                       string validacion = ExcelModifyFunctions.getValueCellString(fila.GetCell(15));

                        // Crear una clave única para la fila


                        if (cellTittleNT.Equals("Número de transacciones:"))
                        {
                            // CASO 1, EL DOCUMENTO SI ES UN FORMATO DE CONSULTA DE MOVIMIENTOS BCO, SIN MODIFICAR
                            return 1;
                        }
                        else if (validacion.Equals("Archivo modificado, Mercantil"))
                        {
                            // CASO 2, EL DOCUMENTO SI ES UN FORMATO DE CONSULTA DE MOVIMIENTOS BCO, MODIFICADO

                            return 2;
                        }
                        else
                        {
                            // CASO 3, EL DOCUMENTO NO ES UN FORMATO DE CONSULTA DE MOVIMIENTOS BCO Mercantil
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



        public void fixFormatOldVersion(TextBox ExcelFilePath)
        {

            ExcelModifyFunctions modifyFunctions = new ExcelModifyFunctions();

            //Guardando la información para que no se dañe al invertir

            List<string> columna1Hoja1 = modifyFunctions.CopyDateColumnsAsStrings(ExcelFilePath.Text, 0, 0);
            List<string> columna2Hoja1 = modifyFunctions.CopyDateColumnsAsStrings(ExcelFilePath.Text, 0, 1);
            List<string> columna3Hoja1 = modifyFunctions.CopyDateColumnsAsStrings(ExcelFilePath.Text, 0, 2);
            List<string> columna4Hoja1 = modifyFunctions.CopyDateColumnsAsStrings(ExcelFilePath.Text, 0, 3);
            List<string> columna5Hoja1 = modifyFunctions.CopyDateColumnsAsStrings(ExcelFilePath.Text, 0, 4);
            List<string> columna6Hoja1 = modifyFunctions.CopyDateColumnsAsStrings(ExcelFilePath.Text, 0, 5);

            List<string> columna1Hoja2 = modifyFunctions.CopyDateColumnsAsStrings(ExcelFilePath.Text, 1, 0);
            List<string> columna2Hoja2 = modifyFunctions.CopyDateColumnsAsStrings(ExcelFilePath.Text, 1, 1);
            List<string> columna3Hoja2 = modifyFunctions.CopyDateColumnsAsStrings(ExcelFilePath.Text, 1, 2);
            List<string> columna4Hoja2 = modifyFunctions.CopyDateColumnsAsStrings(ExcelFilePath.Text, 1, 3);
            List<string> columna5Hoja2 = modifyFunctions.CopyDateColumnsAsStrings(ExcelFilePath.Text, 1, 4);
            List<string> columna6Hoja2 = modifyFunctions.CopyDateColumnsAsStrings(ExcelFilePath.Text, 1, 5);

            //Insertando la información de fechas

            modifyFunctions.changeCellTextFromListInReverseOrderMercantilOldVersion(ExcelFilePath.Text, 0, 0, columna1Hoja1);
            modifyFunctions.changeCellTextFromListInReverseOrderMercantilOldVersion(ExcelFilePath.Text, 1, 0, columna2Hoja1);
            modifyFunctions.changeCellTextFromListInReverseOrderMercantilOldVersion(ExcelFilePath.Text, 2, 0, columna3Hoja1);
            modifyFunctions.changeCellTextFromListInReverseOrderMercantilOldVersion(ExcelFilePath.Text, 3, 0, columna4Hoja1);
            modifyFunctions.changeCellTextFromListInReverseOrderMercantilOldVersion(ExcelFilePath.Text, 4, 0, columna5Hoja1);
            modifyFunctions.changeCellTextFromListInReverseOrderMercantilOldVersion(ExcelFilePath.Text, 5, 0, columna6Hoja1);

            modifyFunctions.changeCellTextFromListInReverseOrderMercantilOldVersion(ExcelFilePath.Text, 0, 1, columna1Hoja2);
            modifyFunctions.changeCellTextFromListInReverseOrderMercantilOldVersion(ExcelFilePath.Text, 1, 1, columna2Hoja2);
            modifyFunctions.changeCellTextFromListInReverseOrderMercantilOldVersion(ExcelFilePath.Text, 2, 1, columna3Hoja2);
            modifyFunctions.changeCellTextFromListInReverseOrderMercantilOldVersion(ExcelFilePath.Text, 3, 1, columna4Hoja2);
            modifyFunctions.changeCellTextFromListInReverseOrderMercantilOldVersion(ExcelFilePath.Text, 4, 1, columna5Hoja2);
            modifyFunctions.changeCellTextFromListInReverseOrderMercantilOldVersion(ExcelFilePath.Text, 5, 1, columna6Hoja2);

            //Cambiando el orden de los movimientos (HOJA 1 Y 2)

            modifyFunctions.CleanColumn(ExcelFilePath.Text, 7, 0);
            modifyFunctions.CleanColumn(ExcelFilePath.Text, 7, 1);
            modifyFunctions.InsertRowOnTop(ExcelFilePath.Text, 0);
            modifyFunctions.InsertRowOnTop(ExcelFilePath.Text, 1);

            //Moviendo columnas para que la función insertar no rompa el orden HOJA 1

            modifyFunctions.MoveColumnsCaseBVnzlaBExterior(ExcelFilePath.Text, 6, 7, 3, 0);
            modifyFunctions.CleanColumn(ExcelFilePath.Text, 6, 0);
            modifyFunctions.MoveColumnsCaseBVnzlaBExterior(ExcelFilePath.Text, 5, 6, 6, 0);
            modifyFunctions.CleanColumn(ExcelFilePath.Text, 5, 0);
            modifyFunctions.MoveColumnsCaseBVnzlaBExterior(ExcelFilePath.Text, 4, 5, 6, 0);
            modifyFunctions.MoveColumnsCaseBVnzlaBExterior(ExcelFilePath.Text, 3, 4, 5, 0);
            modifyFunctions.MoveColumnsCaseBVnzlaBExterior(ExcelFilePath.Text, 2, 3, 5, 0);

            //Moviendo columnas para que la función insertar no rompa el orden HOJA 2

            modifyFunctions.MoveColumnsCaseBVnzlaBExterior(ExcelFilePath.Text, 6, 7, 3, 1);
            modifyFunctions.CleanColumn(ExcelFilePath.Text, 6, 1);
            modifyFunctions.MoveColumnsCaseBVnzlaBExterior(ExcelFilePath.Text, 5, 6, 6, 1);
            modifyFunctions.CleanColumn(ExcelFilePath.Text, 5, 1);
            modifyFunctions.MoveColumnsCaseBVnzlaBExterior(ExcelFilePath.Text, 4, 5, 6, 1);
            modifyFunctions.MoveColumnsCaseBVnzlaBExterior(ExcelFilePath.Text, 3, 4, 5, 1);
            modifyFunctions.MoveColumnsCaseBVnzlaBExterior(ExcelFilePath.Text, 2, 3, 5, 1);

            //Insertando columna de Fecha de validación (HOJA 1 Y 2)

            modifyFunctions.InsertColumnBetweenTwoVersionC3(ExcelFilePath.Text, 2, 0);
            modifyFunctions.InsertColumnBetweenTwoVersionC3(ExcelFilePath.Text, 2, 1);

            //Dando formato a las columnas HOJA 1

            modifyFunctions.FormatNumericColumn(ExcelFilePath.Text, 7, 0);
            modifyFunctions.FormatNumericColumn(ExcelFilePath.Text, 6, 0);
            modifyFunctions.FormatNumericColumn(ExcelFilePath.Text, 5, 0);

            //Dando formato a las columnas HOJA 2

            modifyFunctions.FormatNumericColumn(ExcelFilePath.Text, 7, 1);
            modifyFunctions.FormatNumericColumn(ExcelFilePath.Text, 6, 1);
            modifyFunctions.FormatNumericColumn(ExcelFilePath.Text, 5, 1);

            //Ajustando tamaño de las columnas HOJA 1

            modifyFunctions.AdjustColumnWidth(ExcelFilePath.Text, 3, 15, 0);
            modifyFunctions.AdjustColumnWidth(ExcelFilePath.Text, 4, 45, 0);

            //Ajustando tamaño de las columnas HOJA 2

            modifyFunctions.AdjustColumnWidth(ExcelFilePath.Text, 3, 15, 1);
            modifyFunctions.AdjustColumnWidth(ExcelFilePath.Text, 4, 45, 1);

            //Corrigiendo formato de fecha HOJA 1 Y 2

            ChangeDateFormatCaseMercantil(ExcelFilePath.Text, 1, 0, 0);
            ChangeDateFormatCaseMercantil(ExcelFilePath.Text, 1, 1, 0);

            //Cambiando formato de referencias

            modifyFunctions.ConvertColumnToGeneral(ExcelFilePath.Text, 3, 0);
            modifyFunctions.ConvertColumnToGeneral(ExcelFilePath.Text, 3, 1);

            //Reparando formato de las celdas en blanco (para que no se dañe la fórmula)

            modifyFunctions.replaceEmptyCellsWithZero(ExcelFilePath.Text, 0, 4);
            modifyFunctions.replaceEmptyCellsWithZero(ExcelFilePath.Text, 0, 5);

            modifyFunctions.replaceEmptyCellsWithZero(ExcelFilePath.Text, 1, 4);
            modifyFunctions.replaceEmptyCellsWithZero(ExcelFilePath.Text, 1, 5);

            //Añadiendo identificadores

            modifyFunctions.ChangeCellTextWithFormatAndStyle(ExcelFilePath.Text, 0, 16, "Archivo modificado, Mercantil");
            modifyFunctions.ChangeCellTextWithFormatAndStyle(ExcelFilePath.Text, 0, 17, "Fecha de modificación:" + DateTime.Now.ToString());
            MessageBox.Show("Ajustes realizados exitosamente", "Proceso finalizado", MessageBoxButtons.OK, MessageBoxIcon.Information);





        }


        public void fixFormatNewVersion(TextBox ExcelFilePath)
        {

            ExcelModifyFunctions modifyFunctions = new ExcelModifyFunctions();

            //Guardando la información para que no se dañe al invertir

            List<string> columna1Hoja1 = modifyFunctions.CopyDateColumnsAsStringsMercantil(ExcelFilePath.Text, 0, 0, 1);
            List<string> columna2Hoja1 = modifyFunctions.CopyDateColumnsAsStringsMercantil(ExcelFilePath.Text, 0, 1, 1);
            List<string> columna3Hoja1 = modifyFunctions.CopyDateColumnsAsStringsMercantil(ExcelFilePath.Text, 0, 2, 1);
            List<string> columna4Hoja1 = modifyFunctions.CopyDateColumnsAsStringsMercantil(ExcelFilePath.Text, 0, 3, 1);
            //List<string> columna5Hoja1 = modifyFunctions.CopyDateColumnsAsStringsMercantil(ExcelFilePath.Text, 0, 4, 1);
            //List<string> columna6Hoja1 = modifyFunctions.CopyDateColumnsAsStringsMercantil(ExcelFilePath.Text, 0, 5, 1);

            List<string> columna1Hoja2 = modifyFunctions.CopyDateColumnsAsStringsMercantil(ExcelFilePath.Text, 1, 0, 1);
            List<string> columna2Hoja2 = modifyFunctions.CopyDateColumnsAsStringsMercantil(ExcelFilePath.Text, 1, 1, 1);
            List<string> columna3Hoja2 = modifyFunctions.CopyDateColumnsAsStringsMercantil(ExcelFilePath.Text, 1, 2, 1);
            List<string> columna4Hoja2 = modifyFunctions.CopyDateColumnsAsStringsMercantil(ExcelFilePath.Text, 1, 3, 1);
            //List<string> columna5Hoja2 = modifyFunctions.CopyDateColumnsAsStringsMercantil(ExcelFilePath.Text, 1, 4, 1);
            //List<string> columna6Hoja2 = modifyFunctions.CopyDateColumnsAsStringsMercantil(ExcelFilePath.Text, 1, 5, 1);

            //Insertando la información de fechas

            modifyFunctions.changeCellTextFromListInReverseOrderMercantilNewVersion(ExcelFilePath.Text, 0, 0, columna1Hoja1, 6);
            modifyFunctions.changeCellTextFromListInReverseOrderMercantilNewVersion(ExcelFilePath.Text, 1, 0, columna2Hoja1, 6);
            modifyFunctions.changeCellTextFromListInReverseOrderMercantilNewVersion(ExcelFilePath.Text, 2, 0, columna3Hoja1, 6);
            modifyFunctions.changeCellTextFromListInReverseOrderMercantilNewVersion(ExcelFilePath.Text, 3, 0, columna4Hoja1, 6);

            modifyFunctions.changeCellTextFromListInReverseOrderMercantilNewVersion(ExcelFilePath.Text, 0, 1, columna1Hoja2, 6);
            modifyFunctions.changeCellTextFromListInReverseOrderMercantilNewVersion(ExcelFilePath.Text, 1, 1, columna2Hoja2, 6);
            modifyFunctions.changeCellTextFromListInReverseOrderMercantilNewVersion(ExcelFilePath.Text, 2, 1, columna3Hoja2, 6);
            modifyFunctions.changeCellTextFromListInReverseOrderMercantilNewVersion(ExcelFilePath.Text, 3, 1, columna4Hoja2, 6);


            //Cambiando el orden de los movimientos (HOJA 1 Y 2)

            modifyFunctions.CleanColumn(ExcelFilePath.Text, 7, 0);
            modifyFunctions.CleanColumn(ExcelFilePath.Text, 7, 1);
            modifyFunctions.InsertRowOnTop(ExcelFilePath.Text, 0);
            modifyFunctions.InsertRowOnTop(ExcelFilePath.Text, 1);

            //Moviendo columnas para que la función insertar no rompa el orden HOJA 1

            modifyFunctions.MoveColumnsCaseBVnzlaBExterior(ExcelFilePath.Text, 6, 7, 3, 0);
            modifyFunctions.CleanColumn(ExcelFilePath.Text, 6, 0);
            modifyFunctions.MoveColumnsCaseBVnzlaBExterior(ExcelFilePath.Text, 5, 6, 6, 0);
            modifyFunctions.CleanColumn(ExcelFilePath.Text, 5, 0);
            modifyFunctions.MoveColumnsCaseBVnzlaBExterior(ExcelFilePath.Text, 4, 5, 6, 0);
            modifyFunctions.MoveColumnsCaseBVnzlaBExterior(ExcelFilePath.Text, 3, 4, 5, 0);
            modifyFunctions.MoveColumnsCaseBVnzlaBExterior(ExcelFilePath.Text, 2, 3, 5, 0);

            //Moviendo columnas para que la función insertar no rompa el orden HOJA 2

            modifyFunctions.MoveColumnsCaseBVnzlaBExterior(ExcelFilePath.Text, 6, 7, 3, 1);
            modifyFunctions.CleanColumn(ExcelFilePath.Text, 6, 1);
            modifyFunctions.MoveColumnsCaseBVnzlaBExterior(ExcelFilePath.Text, 5, 6, 6, 1);
            modifyFunctions.CleanColumn(ExcelFilePath.Text, 5, 1);
            modifyFunctions.MoveColumnsCaseBVnzlaBExterior(ExcelFilePath.Text, 4, 5, 6, 1);
            modifyFunctions.MoveColumnsCaseBVnzlaBExterior(ExcelFilePath.Text, 3, 4, 5, 1);
            modifyFunctions.MoveColumnsCaseBVnzlaBExterior(ExcelFilePath.Text, 2, 3, 5, 1);

            //Insertando columna de Fecha de validación (HOJA 1 Y 2)

            modifyFunctions.InsertColumnBetweenTwoVersionC3(ExcelFilePath.Text, 2, 0);
            modifyFunctions.InsertColumnBetweenTwoVersionC3(ExcelFilePath.Text, 2, 1);

            //Separar nùmeros positivos de negativos Hoja 1 y 2

            moveNegativeMovs(ExcelFilePath.Text, 4, 5,0);

            //Dando formato a las columnas HOJA 1

            modifyFunctions.FormatNumericColumn(ExcelFilePath.Text, 6, 0);
            modifyFunctions.FormatNumericColumn(ExcelFilePath.Text, 5, 0);

            //Dando formato a las columnas HOJA 2
             
            modifyFunctions.FormatNumericColumn(ExcelFilePath.Text, 6, 1);
            modifyFunctions.FormatNumericColumn(ExcelFilePath.Text, 5, 1);

            //Corrigiendo formato de fecha HOJA 1 Y 2

            ChangeDateFormatCaseMercantil(ExcelFilePath.Text, 1, 0, 0);
            ChangeDateFormatCaseMercantil(ExcelFilePath.Text, 1, 1, 0);

            //Reparando formato de las celdas en blanco (para que no se dañe la fórmula)

            modifyFunctions.replaceEmptyCellsWithZero(ExcelFilePath.Text, 0, 4);
            modifyFunctions.replaceEmptyCellsWithZero(ExcelFilePath.Text, 0, 5);

            //Ajustando ancho de las celdas Hoja 1 y 2
            modifyFunctions.AdjustColumnWidth(ExcelFilePath.Text, 6, 14, 0);
            modifyFunctions.AdjustColumnWidth(ExcelFilePath.Text, 5, 14, 0);
            modifyFunctions.AdjustColumnWidth(ExcelFilePath.Text, 4, 50, 0);
            modifyFunctions.AdjustColumnWidth(ExcelFilePath.Text, 3, 30, 0);

            modifyFunctions.AdjustColumnWidth(ExcelFilePath.Text, 6, 14, 1);
            modifyFunctions.AdjustColumnWidth(ExcelFilePath.Text, 5, 14, 1);
            modifyFunctions.AdjustColumnWidth(ExcelFilePath.Text, 4, 50, 1);
            modifyFunctions.AdjustColumnWidth(ExcelFilePath.Text, 3, 30, 1);


            //Añadiendo identificadores

            modifyFunctions.ChangeCellTextWithFormatAndStyle(ExcelFilePath.Text, 3, 16, "Archivo modificado, Mercantil");
            modifyFunctions.ChangeCellTextWithFormatAndStyle(ExcelFilePath.Text, 3, 17, "Fecha de modificación:" + DateTime.Now.ToString());
            MessageBox.Show("Ajustes realizados exitosamente", "Proceso finalizado", MessageBoxButtons.OK, MessageBoxIcon.Information);


        }

                 

     public void moveNegativeMovs(string rutaArchivo, int columnaOrigen, int columnaDestino, int nSheet)
        {
            try
            {
                using (FileStream archivo = new FileStream(rutaArchivo, FileMode.Open))
                {
                    IWorkbook libro = new XSSFWorkbook(archivo);
                    string sheetName = libro.GetSheetAt(nSheet).SheetName;
                    ISheet hoja = libro.GetSheet(sheetName);
                    ExcelModifyFunctions modifyFunctions = new ExcelModifyFunctions();

                    // Obtener el estilo de la celda origen (solo una vez)
                    ICell celdaOrigenEjemplo = hoja.GetRow(7).GetCell(4);
                    ICellStyle? estiloOrigen = null;

                    if (celdaOrigenEjemplo != null)
                    {
                        estiloOrigen = celdaOrigenEjemplo.CellStyle;
                        modifyFunctions.CopyCellStyle(celdaOrigenEjemplo.CellStyle, libro);
                    }


                    foreach (IRow fila in hoja)
                    {
                        if (fila != null)
                        {
                            ICell celdaOrigen = fila.GetCell(columnaOrigen);

                            if (celdaOrigen != null && celdaOrigen.CellType == CellType.Numeric)
                            {
                                double valor = celdaOrigen.NumericCellValue;
                                ICell celdaDestino = fila.CreateCell(columnaDestino);
                                // Crear un nuevo estilo para la celda destino
                                ICellStyle estiloDestino = libro.CreateCellStyle();
                                celdaDestino.CellStyle = estiloDestino; 


                                if (valor < 0)
                                {
                                    celdaDestino.SetCellValue(valor * -1);
                                    celdaOrigen.SetCellValue(string.Empty);
                                    
                                }

                            }
                        }
                    }

                    using (FileStream archivoSalida = new FileStream(rutaArchivo, FileMode.Create))
                    {
                        libro.Write(archivoSalida);
                    }
                }

                Console.WriteLine("Cantidades negativas movidas exitosamente.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error al mover cantidades negativas: " + ex.Message);
            }
        }

    }

}


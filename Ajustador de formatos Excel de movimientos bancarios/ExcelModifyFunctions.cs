using System;
using System.Collections.Generic;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.Linq;
using System.Security.Cryptography.Xml;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using NPOI.HSSF.UserModel;
using NPOI.OpenXmlFormats;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;



namespace Ajustador_de_formatos_Excel_de_movimientos_bancarios
{
    internal class ExcelModifyFunctions
    {

        public void AttachExcelFile(ComboBox BankSelector, TextBox ExcelFilePath)
        {

            if (IsOpen(ExcelFilePath.Text))
            {

            }

            else
            {

                if (Path.GetExtension(ExcelFilePath.Text).Equals(".xls"))
                {

                    MessageBox.Show("No se admite ese formato, realiza una conversión al formato XLSX", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

                }

                else
                {

                    if (BankSelector.Text.Equals("Banesco (Modificar)"))
                    {


                        if (BanescoBankValidator(ExcelFilePath.Text) == 1)
                        {


                            // El archivo no está abierto por otro proceso se procede a la ejecución

                            InsertColumnBetweenTwoCaseBanesco(ExcelFilePath.Text, 5);
                            InsertColumnBetweenTwoVersionC2(ExcelFilePath.Text, 2);

                            //Ajustando columna "Fecha de validación"
                            AdjustColumnWidth(ExcelFilePath.Text, 2, 20, 0);
                            //Ajustando columna "Descripción"
                            AdjustColumnWidth(ExcelFilePath.Text, 4, 50, 0);
                            //Ajustando columna "Referencia"
                            AdjustColumnWidth(ExcelFilePath.Text, 3, 14, 0);


                            MoveNegativesNumbersCaseBanesco(ExcelFilePath.Text, 5, 6);

                            //Modificación personal para facilitar la ubicación:
                            /* 
                              * Las columnas se cuentan desde 0, las filas no.
                              * No modifiqué el conteo en las columnas, porque la mayor parte
                              * de las modificaciones a realizar no involucran un conteo demasiado
                              * variado, en el caso de las filas me estaba complicando en las comparaciones.
                              * 
                              * */


                            ChangeCellTextWithFormatAndStyle(ExcelFilePath.Text, 0, 1, "Fecha de validación");
                            ChangeCellTextWithFormatAndStyle(ExcelFilePath.Text, 0, 4, "Ingresos");
                            ChangeCellTextWithFormatAndStyle(ExcelFilePath.Text, 0, 5, "Egresos");
                            ChangeCellTextWithFormatAndStyle(ExcelFilePath.Text, 0, 16, "Archivo modificado");
                            ChangeCellTextWithFormatAndStyle(ExcelFilePath.Text, 0, 17, "Fecha de modificación:" + DateTime.Now.ToString());

                            //Reparando formato de las celdas en blanco (para que no se dañe la fórmula)

                            ReemplazarCeldasEnBlancoConCero(ExcelFilePath.Text, 0, 4);
                            ReemplazarCeldasEnBlancoConCero(ExcelFilePath.Text, 0, 5);
                            ReemplazarCeldasEnBlancoConCero(ExcelFilePath.Text, 0, 6);

              

                            MessageBox.Show("Ajustes realizados exitosamente", "Proceso finalizado", MessageBoxButtons.OK, MessageBoxIcon.Information);


                        }
                        else if (BanescoBankValidator(ExcelFilePath.Text) == 2)
                        {

                            MessageBox.Show("Este archivo ya ha sido modificado", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

                        }
                        else
                        {
                            MessageBox.Show("Error al verificar formato Bco Banesco, verifique el archivo seleccionado", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }


                    }

                    else if (BankSelector.Text.Equals("Banco de Venezuela (Modificar)"))
                    {
                        if (VnzlaBankValidator(ExcelFilePath.Text) == 1)
                        {
                            InsertColumnBetweenTwoCaseBanesco(ExcelFilePath.Text, 5);
                            MoveColumnsCaseBVnzlaBExterior(ExcelFilePath.Text, 6, 5, 3, 0);
                            MoveColumnsCaseBVnzlaBExterior(ExcelFilePath.Text, 4, 6, 3, 0);
                            DeleteColumnAndMove(ExcelFilePath.Text, 4);
                            DeleteColumnAndMove(ExcelFilePath.Text, 6);

                            InsertColumnBetweenTwoVersionC2(ExcelFilePath.Text, 2);
                            ChangeCellTextWithFormatAndStyle(ExcelFilePath.Text, 0, 1, "Fecha de validación");
                            InsertColumnBetweenTwoVersionC2(ExcelFilePath.Text, 6);

                            MoveColumnsCaseBVnzlaBExterior(ExcelFilePath.Text, 5, 6, 1, 0);

                            //Ajustando columna "Fecha de validación"
                            AdjustColumnWidth(ExcelFilePath.Text, 2, 20, 0);
                            //Ajustando columna "Referencias"
                            AdjustColumnWidth(ExcelFilePath.Text, 3, 20, 0);
                            //Ajustando columna "Concepto"
                            AdjustColumnWidth(ExcelFilePath.Text, 4, 50, 0);

                            //Reajustando fuente restantes para mejorar la estética
                            ChangeCellTextWithFormatAndStyle(ExcelFilePath.Text, 0, 0, "Fecha");
                            ChangeCellTextWithFormatAndStyle(ExcelFilePath.Text, 0, 2, "Referencia");
                            ChangeCellTextWithFormatAndStyle(ExcelFilePath.Text, 0, 3, "Concepto");


                            //Borrando las columnas con problemas de formato del banco (no cambian correctamente de General - Número)
                            //y moviendo los datos

                            InsertColumnBetweenTwoVersionC2(ExcelFilePath.Text, 7);
                            InsertColumnBetweenTwoVersionC2(ExcelFilePath.Text, 6);
                            MoveColumnsCaseBVnzlaBExterior(ExcelFilePath.Text, 5, 6, 2, 0);
                            MoveColumnsCaseBVnzlaBExterior(ExcelFilePath.Text, 9, 8, 3, 0);


                            //Ajustando columna "Saldo"
                            AdjustColumnWidth(ExcelFilePath.Text, 7, 15, 0);

                            //Eliminando columnas problemáticas
                            DeleteColumnAndMove(ExcelFilePath.Text, 5);
                            CleanColumn(ExcelFilePath.Text, 8, 0);
                            CleanColumn(ExcelFilePath.Text, 9, 0);
                            CleanColumn(ExcelFilePath.Text, 10, 0);

                            //Ajustando las fuentes y mejorando la estética
                            ChangeCellTextWithFormatAndStyle(ExcelFilePath.Text, 0, 4, "Ingresos");
                            ChangeCellTextWithFormatAndStyle(ExcelFilePath.Text, 0, 5, "Egresos");
                            ChangeCellTextWithFormatAndStyle(ExcelFilePath.Text, 0, 6, "Saldo");


                            //Reparando formato de las celdas en blanco (para que no se dañe la fórmula)

                            ReemplazarCeldasEnBlancoConCero(ExcelFilePath.Text, 0, 4);
                            ReemplazarCeldasEnBlancoConCero(ExcelFilePath.Text, 0, 5);
                            ReemplazarCeldasEnBlancoConCero(ExcelFilePath.Text, 0, 6);

                            MessageBox.Show("Ajustes realizados exitosamente", "Proceso finalizado", MessageBoxButtons.OK, MessageBoxIcon.Information);


                            /*                          
                             * La diferencia entre InsertColumnBetweenTwoCaseBanesco(ExcelFilePath.Text, 5); y 
                             * InsertColumnBetweenTwoCaseBanescoC2(ExcelFilePath.Text, 2); es que la primera inserta
                             * copiando todos los datos de la columna a la izquierda, facilitando el traslado de los datos
                             * mientras que la segunda simplemente inserta una columna en blanco.   
                             *                               
                             * */

                        }
                        else if (VnzlaBankValidator(ExcelFilePath.Text) == 2)
                        {

                            MessageBox.Show("Este archivo ya ha sido modificado", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

                        }
                        else
                        {
                            MessageBox.Show("Error al verificar formato Bco Vnzla, verifique el archivo seleccionado", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }

                    }
                    else if (BankSelector.Text.Equals("Mercantil (Modificar)"))
                    {
                        if (MercantilBankValidator(ExcelFilePath.Text) == 1)
                        {

                            //Guardando la información para que no se dañe al invertir

                            List<string> columna1Hoja1 = CopiarColumnaFechasComoStrings(ExcelFilePath.Text, 0, 0);
                            List<string> columna2Hoja1 = CopiarColumnaFechasComoStrings(ExcelFilePath.Text, 0, 1);
                            List<string> columna3Hoja1 = CopiarColumnaFechasComoStrings(ExcelFilePath.Text, 0, 2);
                            List<string> columna4Hoja1 = CopiarColumnaFechasComoStrings(ExcelFilePath.Text, 0, 3);
                            List<string> columna5Hoja1 = CopiarColumnaFechasComoStrings(ExcelFilePath.Text, 0, 4);
                            List<string> columna6Hoja1 = CopiarColumnaFechasComoStrings(ExcelFilePath.Text, 0, 5);

                            List<string> columna1Hoja2 = CopiarColumnaFechasComoStrings(ExcelFilePath.Text, 1, 0);
                            List<string> columna2Hoja2 = CopiarColumnaFechasComoStrings(ExcelFilePath.Text, 1, 1);
                            List<string> columna3Hoja2 = CopiarColumnaFechasComoStrings(ExcelFilePath.Text, 1, 2);
                            List<string> columna4Hoja2 = CopiarColumnaFechasComoStrings(ExcelFilePath.Text, 1, 3);
                            List<string> columna5Hoja2 = CopiarColumnaFechasComoStrings(ExcelFilePath.Text, 1, 4);
                            List<string> columna6Hoja2 = CopiarColumnaFechasComoStrings(ExcelFilePath.Text, 1, 5);

                            //Insertando la información de fechas

                            //int inverseIndex = columna1Hoja1.Count;
                            //int inverseIndex2 = columna1Hoja2.Count;

                            ChangeCellTextFromListInReverseOrder(ExcelFilePath.Text, 0, 0, columna1Hoja1);
                            ChangeCellTextFromListInReverseOrder(ExcelFilePath.Text, 1, 0, columna2Hoja1);
                            ChangeCellTextFromListInReverseOrder(ExcelFilePath.Text, 2, 0, columna3Hoja1);
                            ChangeCellTextFromListInReverseOrder(ExcelFilePath.Text, 3, 0, columna4Hoja1);
                            ChangeCellTextFromListInReverseOrder(ExcelFilePath.Text, 4, 0, columna5Hoja1);
                            ChangeCellTextFromListInReverseOrder(ExcelFilePath.Text, 5, 0, columna6Hoja1);

                            ChangeCellTextFromListInReverseOrder(ExcelFilePath.Text, 0, 1, columna1Hoja2);
                            ChangeCellTextFromListInReverseOrder(ExcelFilePath.Text, 1, 1, columna2Hoja2);
                            ChangeCellTextFromListInReverseOrder(ExcelFilePath.Text, 2, 1, columna3Hoja2);
                            ChangeCellTextFromListInReverseOrder(ExcelFilePath.Text, 3, 1, columna4Hoja2);
                            ChangeCellTextFromListInReverseOrder(ExcelFilePath.Text, 4, 1, columna5Hoja2);
                            ChangeCellTextFromListInReverseOrder(ExcelFilePath.Text, 5, 1, columna6Hoja2);

                            //for (int i = 0; i < columna1Hoja1.Count; i++)
                            //{
                            //    inverseIndex = inverseIndex - 1;

                            //    ChangeCellText(ExcelFilePath.Text, i, 0, columna1Hoja1[inverseIndex], 0);
                            //    ChangeCellText(ExcelFilePath.Text, i, 1, columna2Hoja1[inverseIndex], 0);
                            //    ChangeCellText(ExcelFilePath.Text, i, 2, columna3Hoja1[inverseIndex], 0);
                            //    ChangeCellText(ExcelFilePath.Text, i, 3, columna4Hoja1[inverseIndex], 0);
                            //    ChangeCellText(ExcelFilePath.Text, i, 4, columna5Hoja1[inverseIndex], 0);
                            //    ChangeCellText(ExcelFilePath.Text, i, 5, columna6Hoja1[inverseIndex], 0);

                            //}

                            //for (int i = 0; i < columna1Hoja2.Count; i++)
                            //{
                            //    inverseIndex2 = inverseIndex2 - 1;

                            //    ChangeCellText(ExcelFilePath.Text, i, 0, columna1Hoja2[inverseIndex2], 1);
                            //    ChangeCellText(ExcelFilePath.Text, i, 1, columna2Hoja2[inverseIndex2], 1);
                            //    ChangeCellText(ExcelFilePath.Text, i, 2, columna3Hoja2[inverseIndex2], 1);
                            //    ChangeCellText(ExcelFilePath.Text, i, 3, columna4Hoja2[inverseIndex2], 1);
                            //    ChangeCellText(ExcelFilePath.Text, i, 4, columna5Hoja2[inverseIndex2], 1);
                            //    ChangeCellText(ExcelFilePath.Text, i, 5, columna6Hoja2[inverseIndex2], 1);

                            //}


                            //Cambiando el orden de los movimientos (HOJA 1 Y 2)

                            CleanColumn(ExcelFilePath.Text, 7, 0);
                            CleanColumn(ExcelFilePath.Text, 7, 1);
                            InsertRowOnTop(ExcelFilePath.Text, 0);
                            InsertRowOnTop(ExcelFilePath.Text, 1);

                            //Moviendo columnas para que la función insertar no rompa el orden HOJA 1

                            MoveColumnsCaseBVnzlaBExterior(ExcelFilePath.Text, 6, 7, 3, 0);
                            CleanColumn(ExcelFilePath.Text, 6, 0);
                            MoveColumnsCaseBVnzlaBExterior(ExcelFilePath.Text, 5, 6, 6, 0);
                            CleanColumn(ExcelFilePath.Text, 5, 0);
                            MoveColumnsCaseBVnzlaBExterior(ExcelFilePath.Text, 4, 5, 6, 0);
                            MoveColumnsCaseBVnzlaBExterior(ExcelFilePath.Text, 3, 4, 5, 0);
                            MoveColumnsCaseBVnzlaBExterior(ExcelFilePath.Text, 2, 3, 5, 0);

                            //Moviendo columnas para que la función insertar no rompa el orden HOJA 2

                            MoveColumnsCaseBVnzlaBExterior(ExcelFilePath.Text, 6, 7, 3, 1);
                            CleanColumn(ExcelFilePath.Text, 6, 1);
                            MoveColumnsCaseBVnzlaBExterior(ExcelFilePath.Text, 5, 6, 6, 1);
                            CleanColumn(ExcelFilePath.Text, 5, 1);
                            MoveColumnsCaseBVnzlaBExterior(ExcelFilePath.Text, 4, 5, 6, 1);
                            MoveColumnsCaseBVnzlaBExterior(ExcelFilePath.Text, 3, 4, 5, 1);
                            MoveColumnsCaseBVnzlaBExterior(ExcelFilePath.Text, 2, 3, 5, 1);

                            //Insertando columna de Fecha de validación (HOJA 1 Y 2)

                            InsertColumnBetweenTwoVersionC3(ExcelFilePath.Text, 2, 0);
                            InsertColumnBetweenTwoVersionC3(ExcelFilePath.Text, 2, 1);


                            //Dando formato a las columnas HOJA 1

                            FormatNumericColumn(ExcelFilePath.Text, 7, 0);
                            FormatNumericColumn(ExcelFilePath.Text, 6, 0);
                            FormatNumericColumn(ExcelFilePath.Text, 5, 0);

                            //Dando formato a las columnas HOJA 2

                            FormatNumericColumn(ExcelFilePath.Text, 7, 1);
                            FormatNumericColumn(ExcelFilePath.Text, 6, 1);
                            FormatNumericColumn(ExcelFilePath.Text, 5, 1);

                            //Ajustando tamaño de las columnas HOJA 1

                            AdjustColumnWidth(ExcelFilePath.Text, 3, 15, 0);
                            AdjustColumnWidth(ExcelFilePath.Text, 4, 45, 0);

                            //Ajustando tamaño de las columnas HOJA 2

                            AdjustColumnWidth(ExcelFilePath.Text, 3, 15, 1);
                            AdjustColumnWidth(ExcelFilePath.Text, 4, 45, 1);

                            //Corrigiendo formato de fecha HOJA 1 Y 2

                            ChangeDateFormatCaseMercantil(ExcelFilePath.Text, 1, 0, 0);
                            ChangeDateFormatCaseMercantil(ExcelFilePath.Text, 1, 1, 0);

                            //Cambiando formato de referencias

                            ConvertColumnToGeneral(ExcelFilePath.Text, 3, 0);
                            ConvertColumnToGeneral(ExcelFilePath.Text, 3, 1);

                            //Reparando formato de las celdas en blanco (para que no se dañe la fórmula)


                            ReemplazarCeldasEnBlancoConCero(ExcelFilePath.Text, 0, 4);
                            ReemplazarCeldasEnBlancoConCero(ExcelFilePath.Text, 0, 5);

                            ReemplazarCeldasEnBlancoConCero(ExcelFilePath.Text, 1, 4);
                            ReemplazarCeldasEnBlancoConCero(ExcelFilePath.Text, 1, 5);

                            //Añadiendo identificadores

                            ChangeCellTextWithFormatAndStyle(ExcelFilePath.Text, 0, 16, "Archivo modificado, Mercantil");
                            ChangeCellTextWithFormatAndStyle(ExcelFilePath.Text, 0, 17, "Fecha de modificación:" + DateTime.Now.ToString());
                            MessageBox.Show("Ajustes realizados exitosamente", "Proceso finalizado", MessageBoxButtons.OK, MessageBoxIcon.Information);

                        }
                        else if (MercantilBankValidator(ExcelFilePath.Text) == 2)
                        {
                            MessageBox.Show("Este archivo ya ha sido modificado", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

                        }
                        else
                        {
                            MessageBox.Show("Error al verificar formato Mercantil, verifique el archivo seleccionado", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

                        }
                     

                    }

                    else if (BankSelector.Text.Equals("Banesco (Ubicar duplicados)"))
                    {


                        ShowDuplicateRows(RevisarFilasRepetidas(ExcelFilePath.Text, 4));


                    }
                    else if (BankSelector.Text.Equals("Exterior (Modificar)"))
                    {
                        if (ExteriorBankValidator(ExcelFilePath.Text) == 1)
                        {

                   
                            List<string> listaColumnaFechas = CopiarColumnaFechasComoStrings(ExcelFilePath.Text, 0, 1);                            


                            //Borrando las columnas con problemas de formato del banco (no cambian correctamente de General - Número)
                            //y moviendo los datos

                            MoveColumnsCaseBVnzlaBExterior(ExcelFilePath.Text, 4, 7, 4, 0);

                            MoveColumnsCaseBVnzlaBExterior(ExcelFilePath.Text, 6, 9, 4, 0);

                            MoveColumnsCaseBVnzlaBExterior(ExcelFilePath.Text, 7, 4, 4, 0);

                            MoveColumnsCaseBVnzlaBExterior(ExcelFilePath.Text, 9, 6, 4, 0);


                            //Limpiando columnas que ya no se utilizarán

                            CleanColumn(ExcelFilePath.Text, 7, 0);
                            CleanColumn(ExcelFilePath.Text, 9, 0);

                            // Modificando las columnas para separar ingresos y egresos y borrar los símbolos + y - que vienen del banco

                             MoveNumberRelatedtoSimbolCaseBancoExterior(ExcelFilePath.Text, 5, 4);

                            //Ajustar Columnas para mejorar la estética
                            AdjustColumnWidth(ExcelFilePath.Text, 1, 15, 0);
                            AdjustColumnWidth(ExcelFilePath.Text, 2, 15, 0);
                            AdjustColumnWidth(ExcelFilePath.Text, 3, 15, 0);
                            AdjustColumnWidth(ExcelFilePath.Text, 4, 35, 0);
                            AdjustColumnWidth(ExcelFilePath.Text, 5, 15, 0);
                            AdjustColumnWidth(ExcelFilePath.Text, 6, 15, 0);
                            AdjustColumnWidth(ExcelFilePath.Text, 7, 15, 0);

                            //Moviendo antes de insertar
                            //Columna total
                            MoveColumnsCaseBVnzlaBExterior(ExcelFilePath.Text, 6, 7, 4, 0);
                            //Columna egresos
                            MoveColumnsCaseBVnzlaBExterior(ExcelFilePath.Text, 5, 6, 2, 0);
                            //Columna ingresos
                            MoveColumnsCaseBVnzlaBExterior(ExcelFilePath.Text, 4, 5, 2, 0);
                            //Columna descripciones
                            MoveColumnsCaseBVnzlaBExterior(ExcelFilePath.Text, 3, 4, 5, 0);
                            //Columna referencias
                            MoveColumnsCaseBVnzlaBExterior(ExcelFilePath.Text, 2, 3, 5, 0);

                            //Dándole formato a las columnas numéricas, punto millar, coma decimal

                            FormatNumericColumn(ExcelFilePath.Text, 5, 0);
                            FormatNumericColumn(ExcelFilePath.Text, 6, 0);
                            FormatNumericColumn(ExcelFilePath.Text, 7, 0);
                            FormatNumericColumn(ExcelFilePath.Text, 8, 0);

                            //Insertando columna de fecha de validación
                            InsertColumnBetweenTwoVersionC3(ExcelFilePath.Text, 2,0);


                            //Cambio de columnas fecha y descripción
                            MoveColumnsCaseBVnzlaBExterior(ExcelFilePath.Text, 1, 8, 5, 0);

                            //Columna referencia
                            MoveColumnsCaseBVnzlaBExterior(ExcelFilePath.Text, 4, 3, 5, 0);
                            //Columna descripción
                            MoveColumnsCaseBVnzlaBExterior(ExcelFilePath.Text, 8, 4, 5, 0);
                            CleanColumn(ExcelFilePath.Text, 8, 0);

                            //Reparando columna de referencia
                            ConvertColumnToGeneral(ExcelFilePath.Text, 3, 0);

                            //Trabajando la columna fecha (formato y posición)
                            //Pegando columna fecha reparada
                            ChangeCellTextFromListInTheSameOrder(ExcelFilePath.Text, 0, 0, listaColumnaFechas);


                            //Ajustando formato de fecha
                            //ChangeDateFormatCaseExterior(ExcelFilePath.Text, 0, 0, 1);
                            //ChangeDateFormatCaseMercantil(ExcelFilePath.Text, 1, 0, 0);
                            ConvertColumnToGeneral(ExcelFilePath.Text, 1, 0);
                            ChangeDateFormatCaseExteriorPrueba(ExcelFilePath.Text, 1, 0, 0);
                            

                            //Ajustando etiqueta de fecha de validación
                            ChangeCellTextWithFormatAndStyle(ExcelFilePath.Text, 0, 1, "Fecha de validación");
                            AdjustColumnWidth(ExcelFilePath.Text, 2, 20, 0);

                            //Marcando la fecha de modificación para validar que el archivo ya fue manipulado
                            ChangeCellTextWithFormatAndStyle(ExcelFilePath.Text, 0, 16, "Archivo modificado, Exterior");
                            ChangeCellTextWithFormatAndStyle(ExcelFilePath.Text, 0, 17, "Fecha de modificación:" + DateTime.Now.ToString());
                                               


                            //Reparando formato de las celdas en blanco (para que no se dañe la fórmula)

                            ReemplazarCeldasEnBlancoConCero(ExcelFilePath.Text, 0, 4);
                            ReemplazarCeldasEnBlancoConCero(ExcelFilePath.Text, 0, 5);
                            ReemplazarCeldasEnBlancoConCero(ExcelFilePath.Text, 0, 6);

                            MessageBox.Show("Ajustes realizados exitosamente", "Proceso finalizado", MessageBoxButtons.OK, MessageBoxIcon.Information);

                        }
                        else if (ExteriorBankValidator(ExcelFilePath.Text) == 2)
                        {

                            MessageBox.Show("Este archivo ya ha sido modificado", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

                        }
                        else
                        {
                            MessageBox.Show("Error al verificar formato Exterior, verifique el archivo seleccionado", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

                        }


                    }
                    else if (BankSelector.Text.Equals("Banco de Vnzla/Exterior (Ubicar duplicados)"))
                    {

                        ShowDuplicateRows(RevisarFilasRepetidas(ExcelFilePath.Text, 3));


                    }
                    else if (BankSelector.Text.Equals("Vzla (Ubicar duplicados - doc general)")){

                        ShowDuplicateRows(LookForDuplicateRowsGeneralDocument(ExcelFilePath.Text, 0, 4));

                    }
                    else if (BankSelector.Text.Equals("Exterior (Ubicar duplicados - doc general)"))
                    {

                        ShowDuplicateRows(LookForDuplicateRowsGeneralDocument(ExcelFilePath.Text,1, 3));

                    }
                    else if (BankSelector.Text.Equals("Mercantil C1 (Ubicar duplicados - doc general)"))
                    {

                        ShowDuplicateRows(LookForDuplicateRowsGeneralDocument(ExcelFilePath.Text, 2, 3));

                    }
                    else if (BankSelector.Text.Equals("Mercantil C2 (Ubicar duplicados - doc general)"))
                    {

                        ShowDuplicateRows(LookForDuplicateRowsGeneralDocument(ExcelFilePath.Text, 3, 3));


                    }
                    else if (BankSelector.Text.Equals("Mercantil C3 (Ubicar duplicados - doc general)"))
                    {

                        ShowDuplicateRows(LookForDuplicateRowsGeneralDocument(ExcelFilePath.Text, 4, 3));

                    }
                    else if (BankSelector.Text.Equals("Banesco (Ubicar duplicados - doc general)"))
                    {

                        ShowDuplicateRows(LookForDuplicateRowsGeneralDocument(ExcelFilePath.Text, 5, 3));

                    }
                    else
                    {
                        MessageBox.Show("Selecciona un formato válido", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                }


            }

        }

        public void InsertColumnBetweenTwoCaseBanesco(string rutaArchivo, int indiceColumna)
        {
            try
            {
                // Abrir el archivo Excel
                FileStream archivo = new FileStream(rutaArchivo, FileMode.Open);
                IWorkbook libro = new XSSFWorkbook(archivo);

                // Obtener la hoja de trabajo
                string nombreHoja = libro.GetSheetAt(0).SheetName;
                
                ISheet hoja = libro.GetSheet(nombreHoja);

                // Desplazar las columnas a la derecha a partir de la columna donde se insertará la nueva
                for (int i = hoja.GetRow(0).LastCellNum; i >= indiceColumna; i--)
                {
                    foreach (IRow fila in hoja)
                    {
                        ICell celdaOrigen = fila.GetCell(i - 1); // Celda de la columna anterior
                        ICell celdaDestino = fila.CreateCell(i); // Nueva celda en la columna actual

                        if (celdaOrigen != null)
                        {
                            // Copiar el valor o fórmula de la celda origen a la celda destino

                            CopyCellValue(celdaOrigen, celdaDestino);

                            // Copiar el estilo de la celda origen a la celda destino (opcional)
                            celdaDestino.CellStyle = celdaOrigen.CellStyle;

                        }
                    }
                }



                // Guardar los cambios
                FileStream archivoSalida = new FileStream(rutaArchivo, FileMode.Create);
                libro.Write(archivoSalida);
                archivoSalida.Close();

                Console.WriteLine("Columna insertada exitosamente.");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }

        }

        public void MoveColumnsCaseBVnzlaBExterior(string rutaArchivo, int columnaOrigen, int columnaDestino, int caseControlerNegativesPositivesorAll, int nHoja)
        {
            try
            {
                using (FileStream archivo = new FileStream(rutaArchivo, FileMode.Open))
                {
                    IWorkbook libro = new XSSFWorkbook(archivo);
                    string nombreHoja = libro.GetSheetAt(nHoja).SheetName;
                    ISheet hoja = libro.GetSheet(nombreHoja);

                    // Obtener el estilo de la celda origen (solo una vez)
                    ICell celdaOrigenEjemplo = hoja.GetRow(0).GetCell(columnaOrigen - 1); // Celda de ejemplo para obtener el estilo
                    ICellStyle? estiloOrigen = null;

                    // Crear un estilo para el formato numérico
                    IDataFormat formatoDatos = libro.CreateDataFormat();
                    IDataFormat formatoCero = libro.CreateDataFormat();

                    if (celdaOrigenEjemplo != null)
                    {
                        estiloOrigen = celdaOrigenEjemplo.CellStyle;
                        CopyCellStyle(celdaOrigenEjemplo.CellStyle, libro);
                    }

                    foreach (IRow fila in hoja)
                    {
                        if (fila != null)
                        {
                            ICell celdaOrigen = fila.GetCell(columnaOrigen - 1);

                            if (celdaOrigen != null && fila.RowNum != 0)
                            {



                                double valor = (double)ObtenerValorCeldaDecimal(celdaOrigen);
                                ICell celdaDestino = fila.CreateCell(columnaDestino - 1);
                                // Crear un nuevo estilo para la celda destino
                                ICellStyle estiloDestino = libro.CreateCellStyle();

                                ICellStyle estiloCero = libro.CreateCellStyle();
                                
                                // Copiar las propiedades del estilo origen al estilo destino
                                if (estiloOrigen != null)
                                {
                                    estiloDestino = CopyCellStyle(estiloOrigen, libro);
                                    estiloDestino.DataFormat = formatoDatos.GetFormat("#,##0.00");
                                }

                                celdaDestino.CellStyle = estiloDestino; // Asignar el nuevo estilo a la celda destino

                                estiloCero.DataFormat = formatoCero.GetFormat("0;-0;;@");

                                //Caso uno, solo los negativos (y los convierte a positivos) y omite todos los positivos originales
                                // dejando la celda de origen en blanco

                                if (celdaOrigen.Equals(""))
                                {
                                    celdaOrigen.SetCellValue(0);
                                    celdaOrigen.CellStyle = estiloCero;
                                }

                                if (caseControlerNegativesPositivesorAll == 1)
                                {
                                    if (valor < 0)
                                    {
                                        celdaDestino.SetCellValue(valor * -1);

                                        celdaOrigen.SetCellValue(0);
                                        celdaOrigen.CellStyle = estiloCero;
                                    }

                                }

                                //Case 2, solo los positivos y omite el 0 EN CORRECCION POR ERRORES EN FORMULAS

                                else if(caseControlerNegativesPositivesorAll == 2)
                                {
                                    if(valor <= 0)
                                    {
                                        celdaDestino.SetCellValue("");
           
                                    }
                                    else
                                    {
                                        celdaDestino.SetCellValue(valor * 1);

                                    }
                                    
                                }

                                //Caso 3, mueve todos los números presentes en la columna

                                else if (caseControlerNegativesPositivesorAll == 3)
                                {
                                    celdaDestino.SetCellValue(valor * 1);
                                    
                                    
                                }
                                //Case 4, especial para Banco Exterior 
                                else if (caseControlerNegativesPositivesorAll == 4)
                                {
                                    string valorString = ObtenerValorCeldaString(celdaOrigen);

                                    if (valorString.Equals('+') || valorString.Equals('-'))
                                    {

                                        celdaDestino.SetCellValue(celdaOrigen.StringCellValue);

                                    }

                                    else
                                    {
                                        celdaDestino.SetCellValue(valor * 1);
                                    }
                                    
                                }
                                //Case 5, solo texto para Banco Mercantil
                                else if (caseControlerNegativesPositivesorAll == 5)
                                {
                                    string valorString = ObtenerValorCeldaString(celdaOrigen);

                                        celdaDestino.SetCellValue(celdaOrigen.StringCellValue);
    
                           

                                }
                                //Case 6, especial para mercantil para Banco Mercantil
                                else if (caseControlerNegativesPositivesorAll == 6)
                                {
                                    if (valor == 0)
                                    {
                                        celdaDestino.SetCellValue("");
  

                                    }
                                    else
                                    {
                                        celdaDestino.SetCellValue(valor * 1);
                                    }

                                }
                                //Case 7, especial para mercantil para Banco Mercantil
                                else if (caseControlerNegativesPositivesorAll == 7)
                                {

                                    DateTime valorDate = ObtenerValorCeldaFecha(celdaOrigen);

                                    celdaDestino.SetCellValue(valorDate);

                                }

                                //Case 8, especial para Exterior, caso primera celda con info distinta a la fecha
                                else if (caseControlerNegativesPositivesorAll == 8)
                                {

                                    if(fila.RowNum == 0)
                                    {

                                        celdaDestino.SetCellValue(celdaOrigen.StringCellValue);

                                    }
                                    else
                                    {
                                        DateTime valorDate = ObtenerValorCeldaFecha(celdaOrigen);

                                        celdaDestino.SetCellValue(valorDate);

                                    }



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


        public void MoveNegativesNumbersCaseBanesco(string rutaArchivo, int columnaOrigen, int columnaDestino)
        {
            try
            {
                using (FileStream archivo = new FileStream(rutaArchivo, FileMode.Open))
                {
                    IWorkbook libro = new XSSFWorkbook(archivo);
                    string nombreHoja = libro.GetSheetAt(0).SheetName;
                    ISheet hoja = libro.GetSheet(nombreHoja);

                    // Obtener el estilo de la celda origen (solo una vez)
                    ICell celdaOrigenEjemplo = hoja.GetRow(0).GetCell(columnaOrigen - 1); // Celda de ejemplo para obtener el estilo
                    ICellStyle? estiloOrigen = null;

                    // Crear un estilo para el formato numérico

                    IDataFormat formatoDatos = libro.CreateDataFormat();


                    if (celdaOrigenEjemplo != null)
                    {
                        estiloOrigen = CopyCellStyle(celdaOrigenEjemplo.CellStyle, libro);
                    }

                    foreach (IRow fila in hoja)
                    {
                        if (fila != null)
                        {
                            ICell celdaOrigen = fila.GetCell(columnaOrigen - 1);

                            if (celdaOrigen != null && celdaOrigen.CellType == CellType.Numeric)
                            {
                                double valor = celdaOrigen.NumericCellValue;
                                ICell celdaDestino = fila.CreateCell(columnaDestino - 1);
                                // Crear un nuevo estilo para la celda destino
                                ICellStyle estiloDestino = libro.CreateCellStyle();

                                ICellStyle estiloCero = libro.CreateCellStyle();

                                
                                estiloCero.DataFormat = libro.CreateDataFormat().GetFormat("0;-0;;@");
                                

                                // Copiar las propiedades del estilo origen al estilo destino
                                if (estiloOrigen != null)
                                {

                                    estiloDestino = CopyCellStyle(estiloOrigen, libro);
                                    estiloDestino.DataFormat = formatoDatos.GetFormat("#,##0.00");


                                }

                                celdaDestino.CellStyle = estiloDestino; // Asignar el nuevo estilo a la celda destino


                                if (valor < 0)
                                {
                                    celdaDestino.SetCellValue(valor * -1);
                                    celdaOrigen.SetCellValue(0);
                                    celdaOrigen.CellStyle = estiloOrigen;
                                    celdaOrigen.CellStyle.DataFormat = estiloCero.DataFormat;
                                }
                                //else if(celdaOrigen == null || celdaOrigen.CellType == CellType.Blank || ObtenerValorCeldaComoString(celdaOrigen).Equals("")){
                                //    celdaOrigen.SetCellValue(0);
                                //    celdaOrigen.CellStyle = estiloOrigen;
                                //    celdaOrigen.CellStyle.DataFormat = estiloCero.DataFormat;
                                //}
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


        public void ChangeCellTextWithFormatAndStyle(string rutaArchivo, int fila, int columna, string texto)
        {
            try
            {
                using (FileStream archivo = new FileStream(rutaArchivo, FileMode.Open))
                {
                    IWorkbook libro = new XSSFWorkbook(archivo);
                    string nombreHoja = libro.GetSheetAt(0).SheetName;
                    ISheet hoja = libro.GetSheet(nombreHoja);

                    IRow filaObj = hoja.GetRow(fila);
                    ICell celda = filaObj.GetCell(columna);

                    if (celda == null)
                    {
                        celda = filaObj.CreateCell(columna - 1);
                    }

                    // 1. Obtener el estilo de la celda (0, 0)
                    ICell celdaEstilo = hoja.GetRow(0).GetCell(0); // Celda de referencia para el estilo
                    ICellStyle estiloBase = null;

                    if (celdaEstilo != null)
                    {
                        estiloBase = celdaEstilo.CellStyle; // Estilo base para copiar formato
                    }


                    // 2. Crear la fuente Calibri 11 en negrita
                    IFont fuente = libro.CreateFont();
                    fuente.FontName = "Calibri";
                    fuente.FontHeightInPoints = 11;
                    fuente.IsBold = true;

                    // 3. Crear un nuevo estilo y asignarle la fuente
                    ICellStyle estilo = libro.CreateCellStyle();
                    

                    // 4. Clonar el formato del delineado del estilo base
                    if (estiloBase != null) // Verifica que el estilo base no sea nulo
                    {
                        estilo = CopyCellStyle(estiloBase, libro);
                        estilo.SetFont(fuente);
                    }

                    // 5. Asignar el estilo a la celda
                    
                    celda.CellStyle = estilo;

                    // 6. Cambiar el texto de la celda
                    celda.SetCellValue(texto);

                    using (FileStream archivoSalida = new FileStream(rutaArchivo, FileMode.Create))
                    {
                        libro.Write(archivoSalida);
                    }
                }

                Console.WriteLine("Texto de celda cambiado, formato y estilo aplicados exitosamente.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error al cambiar texto de celda o aplicar formato/estilo: " + ex.Message);
            }
        }


        public string GetSpecificCellValue(string rutaArchivo, int fila, int columna)
        {
            try
            {
                using (FileStream archivo = new FileStream(rutaArchivo, FileMode.Open))
                {
                    IWorkbook libro = new XSSFWorkbook(archivo);
                    string nombreHoja = libro.GetSheetAt(0).SheetName;
                    ISheet hoja = libro.GetSheet(nombreHoja);

                    // Obtener la fila (los índices comienzan desde 0)
                    IRow filaObj = hoja.GetRow(fila);

                    // Verificar si la fila existe
                    if (filaObj != null)
                    {
                        // Obtener la celda (los índices comienzan desde 0)
                        ICell celda = filaObj.GetCell(columna);

                        // Verificar si la celda existe
                        if (celda != null)
                        {
                            // Obtener el valor de la celda según su tipo de dato
                            switch (celda.CellType)
                            {
                                case CellType.Numeric:
                                    return celda.NumericCellValue.ToString();
                                case CellType.String:
                                    return celda.StringCellValue;
                                case CellType.Formula:
                                    return celda.CellFormula;
                                case CellType.Boolean:
                                    return celda.BooleanCellValue.ToString();
                                case CellType.Error:
                                    return celda.ErrorCellValue.ToString();
                                case CellType.Blank:
                                    return "";
                                default:
                                    return "Tipo de dato no manejado: " + celda.CellType;
                            }
                        }
                        else
                        {
                            return "La celda no existe.";
                        }
                    }
                    else
                    {
                        return "La fila no existe.";
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error al obtener valor de la celda: " + ex.Message);
                return null;
            }
        }


        public void InsertColumnBetweenTwoVersionC2(string rutaArchivo, int indiceColumna)
        {
            try
            {
                // Abrir el archivo Excel
                FileStream archivo = new FileStream(rutaArchivo, FileMode.Open);
                IWorkbook libro = new XSSFWorkbook(archivo);

                // Obtener la hoja de trabajo
                string nombreHoja = libro.GetSheetAt(0).SheetName;

                ISheet hoja = libro.GetSheet(nombreHoja);

                // Desplazar las columnas a la derecha a partir de la columna donde se insertará la nueva
                for (int i = hoja.GetRow(0).LastCellNum; i >= indiceColumna; i--)
                {
                    foreach (IRow fila in hoja)
                    {
                        ICell celdaOrigen = fila.GetCell(i - 1); // Celda de la columna anterior
                        ICell celdaDestino = fila.CreateCell(i); // Nueva celda en la columna actual

                        if (celdaOrigen != null)
                        {
                            // Copiar el valor o fórmula de la celda origen a la celda destino

                            CopyCellValue(celdaOrigen, celdaDestino);

                            // Copiar el estilo de la celda origen a la celda destino (opcional)
                            celdaDestino.CellStyle = celdaOrigen.CellStyle;
                            celdaOrigen.SetCellValue("");

                        }
                    }
                }



                // Guardar los cambios
                FileStream archivoSalida = new FileStream(rutaArchivo, FileMode.Create);
                libro.Write(archivoSalida);
                archivoSalida.Close();

                Console.WriteLine("Columna insertada exitosamente.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error al insertar columna: " + ex.Message);
            }
        }

        public void InsertColumnBetweenTwoVersionC3(string rutaArchivo, int indiceColumna, int nHoja)
        {
            try
            {
                using (FileStream archivo = new FileStream(rutaArchivo, FileMode.Open))
                {
                    IWorkbook libro = new XSSFWorkbook(archivo);
                    string nombreHoja = libro.GetSheetAt(nHoja).SheetName;
                    ISheet hoja = libro.GetSheet(nombreHoja);

                    // Encontrar la última columna con contenido antes de la columna insertada
                    int ultimaColumnaConContenido = 0;
                    foreach (IRow fila in hoja)
                    {
                        if (fila != null)
                        {
                            for (int i = indiceColumna - 2; i >= 0; i--)
                            {
                                if (fila.GetCell(i) != null)
                                {
                                    ultimaColumnaConContenido = Math.Max(ultimaColumnaConContenido, indiceColumna);
                                    break;
                                }
                            }
                        }
                    }

                    // Desplazar las columnas a la derecha solo hasta la última columna con contenido
                    for (int i = ultimaColumnaConContenido; i >= indiceColumna; i--)
                    {
                        foreach (IRow fila in hoja)
                        {
                            if (fila != null)
                            {
                                ICell celdaOrigen = fila.GetCell(i - 1);
                                ICell celdaDestino = fila.CreateCell(i);

                                if (celdaOrigen != null)
                                {
                                    CopyCellValue(celdaOrigen, celdaDestino);
                                    celdaDestino.CellStyle = celdaOrigen.CellStyle;
                                    celdaOrigen.SetCellValue("");
                                }
                            }
                        }
                    }

                    using (FileStream archivoSalida = new FileStream(rutaArchivo, FileMode.Create))
                    {
                        libro.Write(archivoSalida);
                    }

                    Console.WriteLine("Columna insertada exitosamente.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error al insertar columna: " + ex.Message);
            }
        }

        public void AdjustColumnWidth(string rutaArchivo, int columna, double ancho, int nHoja)
        {
            try
            {
                using (FileStream archivo = new FileStream(rutaArchivo, FileMode.Open))
                {
                    IWorkbook libro = null;

                    // Determinar el tipo de archivo y abrir el libro
                    if (rutaArchivo.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
                    {
                        libro = new XSSFWorkbook(archivo);
                    }
                    else if (rutaArchivo.EndsWith(".xls", StringComparison.OrdinalIgnoreCase))
                    {
                        libro = new HSSFWorkbook(archivo);
                    }
                    else
                    {
                        throw new Exception("Formato de archivo no compatible. Debe ser .xlsx o .xls.");
                    }

                    // Obtener la hoja de trabajo
                    string nombreHoja = libro.GetSheetAt(nHoja).SheetName;
                    ISheet hoja = libro.GetSheet(nombreHoja);

                    // Ajustar el ancho de la columna
                    hoja.SetColumnWidth(columna - 1, (int)(ancho * 256)); // Multiplicar por 256 para obtener el ancho en unidades de Excel

                    using (FileStream archivoSalida = new FileStream(rutaArchivo, FileMode.Create))
                    {
                        libro.Write(archivoSalida);
                    }
                }

                Console.WriteLine("Ancho de columna ajustado exitosamente.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error al ajustar ancho de columna: " + ex.Message);
            }
        }

        public bool IsOpen(string path)
        {
            try
            {
                using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.None))
                {
                    // El archivo no está abierto por otro proceso
                    return false;
                }
            }
            catch (IOException ex)
            {
                // El archivo está abierto por otro proceso
                if (IsFileLocked(ex))
                {
                    MessageBox.Show("El archivo está abierto. Cierre el archivo para continuar.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    // Otro error de E/S
                    MessageBox.Show("Error al acceder al archivo: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return true;
            }
        }


        // Función auxiliar para verificar si la excepción es por archivo bloqueado
        public virtual bool IsFileLocked(IOException e)
        {
            return e.HResult == -2147024864; // Error HRESULT: 0x80070020
        }

      
     
        public void ShowListofText(List<string> lineas)
        {
            try
            {
                // Verificar si la lista de líneas está vacía
                if (lineas == null || lineas.Count == 0)
                {
                    MessageBox.Show("No hay texto para mostrar.", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // Construir el texto a mostrar
                string texto = string.Join(Environment.NewLine, lineas);

                // Mostrar el texto en una ventana informativa
                MessageBox.Show(texto, "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al mostrar texto: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        public List<string> RevisarFilasRepetidas(string rutaArchivo, int startedRowToRevision)
        {
            List<string> lineas = new List<string>();
            List<string> lineasToHL = new List<string>();
            try
            {
                using (FileStream archivo = new FileStream(rutaArchivo, FileMode.Open))
                {
                    IWorkbook libro = new XSSFWorkbook(archivo);
                    string nombreHoja = libro.GetSheetAt(0).SheetName;
                    ISheet hoja = libro.GetSheet(nombreHoja);
                    

                    Dictionary<string, int> filasUnicas = new Dictionary<string, int>();

                    for (int i = startedRowToRevision; i <= hoja.LastRowNum; i++)
                    {
                        IRow fila = hoja.GetRow(i);
                        if (fila != null)
                        {
                            DateTime fecha = ObtenerValorCeldaFecha(fila.GetCell(0));
                            DateTime fechaValidacion = ObtenerValorCeldaFecha(fila.GetCell(1));
                            string referencia = ObtenerValorCeldaString(fila.GetCell(2)).Trim().ToLower(); // Normalizar referencia
                            string descripcion = ObtenerValorCeldaString(fila.GetCell(3)).Trim().ToLower(); // Normalizar descripción
                            decimal ingresos = ObtenerValorCeldaDecimal(fila.GetCell(4));
                            decimal egresos = ObtenerValorCeldaDecimal(fila.GetCell(5));

                            // Crear una clave única para la fila
                            string filaHash = $"{fecha.Date:dd-MM-yyyy}{referencia}{descripcion}{ingresos} {egresos}";

                            if (filasUnicas.ContainsKey(filaHash))
                            {                                
                                
                                HighlightRow(rutaArchivo, i, 0);
                                lineas.Add($"Fila repetida encontrada en la fila {i + 1}. \n \n" +
                                    $"Fecha: {fecha}, Fecha Validación: {fechaValidacion}, Referencia: {referencia}, Descripción: {descripcion}, Ingresos: {ingresos}, Egresos: {egresos} \n \n");
                                                                
                                // Para depurar
                                //lineas.Add($"Fila data: {fecha} {fechaValidacion} {referencia} {descripcion} {ingresos} {egresos}");

                               
                            }
                            else
                            {
                                filasUnicas.Add(filaHash, i + 1);
                            }
                        }
                    }

                    if (lineas.Count == 0)
                    {
                        lineas.Add("No se encontraron filas repetidas.");
                    }

                    
                    
                    
                }
            }
            catch (Exception ex)
            {
                lineas.Add("Error al revisar filas repetidas: " + ex.Message);
            }
            return lineas;
        }

        public List<string> LookForDuplicateRowsGeneralDocument(string rutaArchivo, int NSheet, int startedRowToRevision)
        {
            List<string> lineas = new List<string>();
            try
            {
                using (FileStream archivo = new FileStream(rutaArchivo, FileMode.Open))
                {
                    IWorkbook libro = new XSSFWorkbook(archivo);
                    string nombreHoja = libro.GetSheetAt(NSheet).SheetName;
                    ISheet hoja = libro.GetSheet(nombreHoja);

                    Dictionary<string, int> filasUnicas = new Dictionary<string, int>();

                    for (int i = startedRowToRevision; i <= hoja.LastRowNum; i++)
                    {
                        IRow fila = hoja.GetRow(i);
                        if (fila != null)
                        {
                            string fecha = ObtenerValorCeldaString(fila.GetCell(0));
                            string referencia = ObtenerValorCeldaString(fila.GetCell(1));
                            string descripcion = ObtenerValorCeldaString(fila.GetCell(2)).Trim().ToLower(); // Normalizar referencia
                            decimal ingresos = ObtenerValorCeldaDecimal(fila.GetCell(3));
                            decimal egresos = ObtenerValorCeldaDecimal(fila.GetCell(4));

                            // Crear una clave única para la fila
                            string filaHash = $"{fecha}{referencia}{descripcion}{ingresos} {egresos}";

                            if (filasUnicas.ContainsKey(filaHash))
                            {

                                lineas.Add($"Fila repetida encontrada en la fila {i + 1}. \n \n" +
                                    $"Fecha: {fecha}, Referencia: {referencia}, Descripción: {descripcion}, Ingresos: {ingresos}, Egresos: {egresos} \n \n");

                                // Para depurar
                                //lineas.Add($"Fila data: {fecha} {fechaValidacion} {referencia} {descripcion} {ingresos} {egresos}");

                            }
                            else
                            {
                                filasUnicas.Add(filaHash, i + 1);
                            }
                        }
                    }
                    if (lineas.Count == 0)
                    {
                        lineas.Add("No se encontraron filas repetidas.");
                    }


                }
            }
            catch (Exception ex)
            {
                lineas.Add("Error al revisar filas repetidas: " + ex.Message);
            }
            return lineas;
        }

        // Funciones auxiliares para obtener los valores de las celdas
        public DateTime ObtenerValorCeldaFecha(ICell celda)
        {
            if (celda != null && celda.CellType == CellType.Numeric && DateUtil.IsCellDateFormatted(celda))
            {
                return DateUtil.GetJavaDate(celda.NumericCellValue);
            }
            return DateTime.MinValue;
        }

        public string ObtenerValorCeldaString(ICell celda)
        {
            if (celda != null)
            {
                switch (celda.CellType)
                {
                    case CellType.String:
                        return celda.StringCellValue;
                    case CellType.Numeric:
                        return celda.NumericCellValue.ToString();
                    case CellType.Formula:
                        return celda.CellFormula;
                    default:
                        return celda.ToString();
                }
            }
            return "";
        }

        public decimal ObtenerValorCeldaDecimal(ICell celda)
        {
            if (celda != null && celda.CellType == CellType.Numeric)
            {
                return (decimal)celda.NumericCellValue;
            }
            if (celda == null || celda.CellType == CellType.Blank)
            {
                return 0;
            }
            if (celda.CellType == CellType.String)
            {
                decimal valor;
                if (decimal.TryParse(celda.StringCellValue.Trim(), out valor))
                {
                    return valor;
                }
                else
                {
                    return 0;
                }
            }
            return 0;
        }


        public void DeleteColumnAndMove(string rutaArchivo, int columna)
        {
            try
            {
                using (FileStream archivo = new FileStream(rutaArchivo, FileMode.Open))
                {
                    IWorkbook libro = new XSSFWorkbook(archivo);
                    string nombreHoja = libro.GetSheetAt(0).SheetName;
                    ISheet hoja = libro.GetSheet(nombreHoja);

                    for (int filaIndex = 0; filaIndex <= hoja.LastRowNum; filaIndex++)
                    {
                        IRow fila = hoja.GetRow(filaIndex);
                        if (fila != null)
                        {
                            fila.RemoveCell(fila.GetCell(columna - 1));

                            for (int celdaIndex = columna; celdaIndex <= fila.LastCellNum; celdaIndex++)
                            {
                                ICell celda = fila.GetCell(celdaIndex);
                                if (celda != null)
                                {
                                    ICell nuevaCelda = fila.CreateCell(celdaIndex - 1, celda.CellType);
                                    CopyCellValue(celda, nuevaCelda);
                                    nuevaCelda.CellStyle = CopyCellStyle(celda.CellStyle, libro);
                                    
                                }
                            }
                            if (fila.LastCellNum > 0)
                            {
                                fila.RemoveCell(fila.GetCell(fila.LastCellNum - 1));
                            }
                        }
                    }

                    using (FileStream archivoSalida = new FileStream(rutaArchivo, FileMode.Create))
                    {
                        libro.Write(archivoSalida);
                    }
                }
                Console.WriteLine($"Columna {columna} eliminada exitosamente.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error al eliminar columna: {ex.Message}");
            }
        }


         private void CopyCellValue(ICell origen, ICell destino)
        {
            switch (origen.CellType)
            {
                case CellType.Boolean:
                    destino.SetCellValue(origen.BooleanCellValue);
                    break;
                case CellType.Numeric:
                        destino.SetCellValue(origen.NumericCellValue);
                    
                    break;
                case CellType.String:
                    destino.SetCellValue(origen.StringCellValue);
                    break;
                case CellType.Formula:
                    destino.SetCellFormula(origen.CellFormula);
                    break;
                case CellType.Blank:
                    destino.SetCellValue("");
                    break;
                case CellType.Error:
                    destino.SetCellErrorValue(origen.ErrorCellValue);
                    break;
                default:
                    break;
            }
        }


        

        // Función auxiliar para copiar el contenido y formato de una celd

        private ICellStyle CopyCellStyle(ICellStyle estiloOrigen, IWorkbook libro)
        {

            ICellStyle estiloDestino = libro.CreateCellStyle();

            estiloDestino.Alignment = estiloOrigen.Alignment;
            estiloDestino.BorderBottom = estiloOrigen.BorderBottom;
            estiloDestino.BorderLeft = estiloOrigen.BorderLeft;
            estiloDestino.BorderRight = estiloOrigen.BorderRight;
            estiloDestino.BorderTop = estiloOrigen.BorderTop;
            estiloDestino.DataFormat = estiloOrigen.DataFormat;
            estiloDestino.FillBackgroundColor = estiloOrigen.FillBackgroundColor;
            estiloDestino.FillForegroundColor = estiloOrigen.FillForegroundColor;
            estiloDestino.FillPattern = estiloOrigen.FillPattern;

            return estiloDestino;

        }

        public void CleanColumn(string rutaArchivo, int columna, int indiceHoja)
        {
            try
            {
                using (FileStream archivo = new FileStream(rutaArchivo, FileMode.Open))
                {
                    IWorkbook libro = new XSSFWorkbook(archivo);
                    ISheet hoja = null;

                    try
                    {
                        hoja = libro.GetSheetAt(indiceHoja);
                    }
                    catch (ArgumentException ex)
                    {
                        Console.WriteLine($"Error: No se encontró la hoja en el índice {indiceHoja}. {ex.Message}");
                        return; // Salir de la función si no se encuentra la hoja
                    }

                    if (hoja != null)
                    {
                        // La hoja existe, puedes realizar las operaciones que necesites
                        string nombreHoja = hoja.SheetName;
                        Console.WriteLine($"Nombre de la hoja: {nombreHoja}");

                        // ... (resto de tu código) ...
                        for (int filaIndex = 0; filaIndex <= hoja.LastRowNum; filaIndex++)
                        {
                            IRow fila = hoja.GetRow(filaIndex);
                            if (fila != null)
                            {
                                ICell celda = fila.GetCell(columna - 1);
                                if (celda != null)
                                {
                                    // Poner la celda en blanco
                                    celda.SetCellValue("");

                                    // Eliminar estilos de borde y similares
                                    ICellStyle estilo = libro.CreateCellStyle();
                                    celda.CellStyle = estilo;
                                }
                            }
                        }

                        using (FileStream archivoSalida = new FileStream(rutaArchivo, FileMode.Create))
                        {
                            libro.Write(archivoSalida);
                        }
                    }
                    else
                    {
                        // La hoja no existe, no se ejecuta nada
                        Console.WriteLine($"La hoja no existe en el índice {indiceHoja}.");
                    }

                    Console.WriteLine($"Columna {columna} limpiada exitosamente.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error al limpiar columna: {ex.Message}");
            }
        }

        private int VnzlaBankValidator(string rutaArchivo)
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
                            string fecha = ObtenerValorCeldaString(fila.GetCell(0));
                            string fechaValidacion = ObtenerValorCeldaString(fila.GetCell(1)).Trim().ToLower();
                            string referencia = ObtenerValorCeldaString(fila.GetCell(1)).Trim().ToLower(); // Normalizar referencia
                            string concepto = ObtenerValorCeldaString(fila.GetCell(2)).Trim().ToLower(); // Normalizar descripción
                            string saldo = ObtenerValorCeldaString(fila.GetCell(3)).Trim().ToLower();
                            string monto = ObtenerValorCeldaString(fila.GetCell(4)).Trim().ToLower();
                            string tipoMov = ObtenerValorCeldaString(fila.GetCell(5)).Trim().ToLower();
                            string rif = ObtenerValorCeldaString(fila.GetCell(6)).Trim().ToLower();
                            string numeroCuenta = ObtenerValorCeldaString(fila.GetCell(7)).Trim().ToLower();


                            // Crear una clave única para la fila
                            string filaHash = $"{fecha}{referencia}{concepto}{saldo}{monto}{tipoMov}{rif}{numeroCuenta}";

                        if (filaHash.Equals("fechareferenciaconceptosaldomontotipomovimientorifnumerocuenta")){
                            // CASO 1, EL DOCUMENTO SI ES UN FORMATO DE CONSULTA DE MOVIMIENTOS BCO VNZLA, SIN MODIFICAR
                            return 1;
                        }else if (fechaValidacion.Equals("fecha de validación")){
                            // CASO 2, EL DOCUMENTO SI ES UN FORMATO DE CONSULTA DE MOVIMIENTOS BCO VNZLA, MODIFICADO
                            
                            return 2;
                        }
                        else
                        {
                            // CASO 3, EL DOCUMENTO NO ES UN FORMAATO DE CONSULTA DE MOVIMIENTOS BCO DE VNZLA
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

        private int BanescoBankValidator(string rutaArchivo)
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
                        string fecha = ObtenerValorCeldaString(fila.GetCell(0)).Trim().ToLower();
                        string fechaValidacion = ObtenerValorCeldaString(fila.GetCell(1)).Trim().ToLower();
                        string referencia = ObtenerValorCeldaString(fila.GetCell(1)).Trim().ToLower(); // Normalizar referencia
                        string descripcion = ObtenerValorCeldaString(fila.GetCell(2)).Trim().ToLower(); // Normalizar descripción
                        string monto = ObtenerValorCeldaString(fila.GetCell(3)).Trim().ToLower();
                        string balance = ObtenerValorCeldaString(fila.GetCell(4)).Trim().ToLower();


                        // Crear una clave única para la fila
                        string filaHash = $"{fecha}{referencia}{descripcion}{monto}{balance}";
                        

                        if (filaHash.Equals("fechareferenciadescripciónmontobalance"))
                        {
                            // CASO 1, EL DOCUMENTO SI ES UN FORMATO DE CONSULTA DE MOVIMIENTOS BCO, SIN MODIFICAR
                            return 1;
                        }
                        else if (fechaValidacion.Equals("fecha de validación"))
                        {
                            // CASO 2, EL DOCUMENTO SI ES UN FORMATO DE CONSULTA DE MOVIMIENTOS BCO, MODIFICADO

                            return 2;
                        }
                        else
                        {
                            // CASO 3, EL DOCUMENTO NO ES UN FORMATO DE CONSULTA DE MOVIMIENTOS BCO BANESCO
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

        private int MercantilBankValidator(string rutaArchivo)
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

                        string validacion = ObtenerValorCeldaString(fila.GetCell(15));    
                        string descripcion = ObtenerValorCeldaString(fila.GetCell(6)); // Normalizar descripción


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

        private int ExteriorBankValidator(string rutaArchivo)
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
                        string nCuenta = ObtenerValorCeldaString(fila.GetCell(1));
                        string validacion = ObtenerValorCeldaString(fila.GetCell(15));
       


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


        private void ShowDuplicateRows(List<string> lineas)
        {

            try
            {
                for (int i = 0; i < lineas.Count; i++) {
                    // Verificar si la lista de líneas está vacía
                    if (lineas == null || lineas.Count == 0)
                    {
                        MessageBox.Show("No hay texto para mostrar.", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        break;
                    }


                    // Mostrar el texto en una ventana informativa
                    DialogResult resultado = MessageBox.Show(lineas[i] + "\n" + "Presiona Ok, para continuar la búsqueda una vez hayas eliminado el duplicado o cancelar si deseas abortar la busqueda.", "Información", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                  
                    // MessageBox.Show(lineas[i] + "\n" + "Presiona Ok, para continuar la búsqueda una vez hayas eliminado el duplicado", "Información", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);

                    if (resultado == DialogResult.OK)
                    {
                        // El usuario presionó "Aceptar"
                        // Continúa con el bucle
                            
                    }
                    else
                    {
                        // El usuario presionó "Cancelar"
                        // Rompe el bucle
                        break;
                        
                    }



                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al mostrar texto: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

   
        }

        /*
         * 
         * Esta función permite mover los números negativos de una columna, en función del signo contenido en su
         * correlativo en la columna justo a su derecha. Fue usada para ajustar el formato de banco exterior, y que 
         * las cantidades positivas y negativas no estén juntas en una columnas, siendo diferenciadas solo por esos
         * símbolos positivos y negativos en la columna a la derecha. 
         * 
         * 
         * */
        public void MoveNumberRelatedtoSimbolCaseBancoExterior(string rutaArchivo,  int columnaSimbolos, int columnaNumeros)
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
                        CopyCellStyle(celdaOrigenEjemplo.CellStyle, libro);
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

        //public void ReverseColumns(string rutaArchivo, int sheetName)
        //{
        //    try
        //    {
        //        using (FileStream archivo = new FileStream(rutaArchivo, FileMode.Open))
        //        {
        //            IWorkbook libro = new XSSFWorkbook(archivo);

        //            try
        //            {
        //                string nombreHoja = libro.GetSheetAt(sheetName).SheetName;
        //                ISheet hoja = libro.GetSheet(nombreHoja);
        //                int ultimaFila = hoja.LastRowNum;

        //                //MessageBox.Show(ultimaFila.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //                List<IRow> filas = new List<IRow>();

        //                // Almacenar todas las filas en una lista
        //                for (int i = 0; i <= ultimaFila; i++)
        //                {
        //                    IRow fila = hoja.GetRow(i);
        //                    if (fila != null)
        //                    {
        //                        filas.Add(fila);

        //                    }
        //                }

        //                MessageBox.Show(filas.Count.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //                // Reinsertar las filas en orden inverso
        //                int nuevaFilaIndex = 0;
        //                for (int i = filas.Count - 1; i >= 0; i--)
        //                {
        //                    IRow nuevaFila = hoja.CreateRow(nuevaFilaIndex);
        //                    CopyRow(filas[i], nuevaFila);
        //                    nuevaFilaIndex++;

        //                    Console.WriteLine($"Fila {i} reinsertada en {nuevaFilaIndex - 1}."); // Depuración
        //                }


        //                // Guardar los cambios
        //                using (FileStream archivoSalida = new FileStream(rutaArchivo, FileMode.Create))
        //                {
        //                    libro.Write(archivoSalida);
        //                }
        //            }
        //            catch (ArgumentException ex)
        //            {
        //                Console.WriteLine($"Error: No se encontró la hoja en el índice {sheetName}. {ex.Message}");
        //                return; // Salir de la función si no se encuentra la hoja
        //            }



        //        }
        //        Console.WriteLine($"Orden de filas invertido exitosamente en la hoja.");
        //    }
        //    catch (Exception ex)
        //    {
        //        Console.WriteLine($"Error al invertir orden de filas: {ex.Message}");
        //    }
        //}


        public void ReverseColumns(string rutaArchivo, int sheetName)
        {
            try
            {
                using (FileStream archivo = new FileStream(rutaArchivo, FileMode.Open))
                {
                    IWorkbook libro = new XSSFWorkbook(archivo);

                    try
                    {
                        string nombreHoja = libro.GetSheetAt(sheetName).SheetName;
                        ISheet hoja = libro.GetSheet(nombreHoja);
                        int ultimaFila = hoja.PhysicalNumberOfRows;

                        List<IRow> filas = new List<IRow>();

                        // Almacenar todas las filas en una lista
                        for (int i = 0; i < ultimaFila; i++)
                        {
                            IRow fila = hoja.GetRow(i);
                            if (fila != null)
                            {
                                filas.Add(fila);
                            }
                        }

                        // Reinsertar las filas en orden inverso, sin eliminar filas
                        int nuevaFilaIndex = 0;
                        for (int i = filas.Count - 1; i >= 0; i--)
                        {
                            IRow nuevaFila = hoja.CreateRow(nuevaFilaIndex);
                            CopyRow(filas[i], nuevaFila);
                            nuevaFilaIndex++;
                            Console.WriteLine($"Fila {i} reinsertada en {nuevaFilaIndex - 1}."); // Depuración
                        }

                        // Guardar los cambios
                        using (FileStream archivoSalida = new FileStream(rutaArchivo, FileMode.Create))
                        {
                            libro.Write(archivoSalida);
                        }
                    }
                    catch (ArgumentException ex)
                    {
                        Console.WriteLine($"Error: No se encontró la hoja en el índice {sheetName}. {ex.Message}");
                        return; // Salir de la función si no se encuentra la hoja
                    }
                }
                Console.WriteLine($"Orden de filas invertido exitosamente en la hoja.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error al invertir orden de filas: {ex.Message}");
            }
        }

        private void CopyRow(IRow filaOriginal, IRow filaNueva)
        {
            if (filaOriginal == null) return;

            for (int i = 0; i < filaOriginal.Cells.Count; i++)
            {
                ICell celdaOriginal = filaOriginal.GetCell(i);
                ICell celdaNueva = filaNueva.CreateCell(i);

                if (celdaOriginal != null)
                {
                    // Copiar valor de la celda
                    celdaNueva.SetCellValue(celdaOriginal.ToString());

                    // Opcional: Copiar estilos, si es necesario
                    // celdaNueva.CellStyle = celdaOriginal.CellStyle;  // Si quieres copiar el estilo también.
                }
            }
        }


        public void FormatNumericColumn(string rutaArchivo, int columna, int nHoja)
        {
            try
            {
                using (FileStream archivo = new FileStream(rutaArchivo, FileMode.Open))
                {
                    IWorkbook libro = new XSSFWorkbook(archivo);
                    string nombreHoja = libro.GetSheetAt(nHoja).SheetName;
                    ISheet hoja = libro.GetSheet(nombreHoja);

                    // Crear un estilo para el formato numérico
                    ICellStyle estiloNumero = libro.CreateCellStyle();
                    IDataFormat formatoDatos = libro.CreateDataFormat();
                    estiloNumero.DataFormat = formatoDatos.GetFormat("#,##0.00");

                    ICellStyle estiloCero = libro.CreateCellStyle();
                    estiloCero.DataFormat = libro.CreateDataFormat().GetFormat("0;-0;;@");

                    for (int filaIndex = 0; filaIndex <= hoja.LastRowNum; filaIndex++)
                    {
                        IRow fila = hoja.GetRow(filaIndex);
                        if (fila != null)
                        {
                            ICell celda = fila.GetCell(columna - 1);
                            if (celda != null && celda.CellType == CellType.Numeric)
                            {
                                celda.CellStyle = estiloNumero;
                                //Depurando (If)
                                if (celda.Equals(""))
                                {
                                    celda.SetCellValue(0);
                                    celda.CellStyle = estiloCero;
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
                Console.WriteLine($"Columna {columna} formateada exitosamente.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error al formatear columna: {ex.Message}");
            }
        }

          public void InsertRowOnTop(string rutaArchivo, int nHoja)
        {
            try
            {
                using (FileStream archivo = new FileStream(rutaArchivo, FileMode.Open))
                {
                    IWorkbook libro = new XSSFWorkbook(archivo);
                    string nombreHoja = libro.GetSheetAt(nHoja).SheetName;
                    ISheet hoja = libro.GetSheet(nombreHoja);

                    // Desplazar las filas existentes hacia abajo
                    hoja.ShiftRows(0, hoja.LastRowNum, 1, true, true);

                    // Crear una nueva fila en la parte superior
                    IRow nuevaFila = hoja.CreateRow(0);

                    // Guardar los cambios
                    using (FileStream archivoSalida = new FileStream(rutaArchivo, FileMode.Create))
                    {
                        libro.Write(archivoSalida);
                    }
                }
                Console.WriteLine("Fila insertada en la parte superior exitosamente.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error al insertar fila: {ex.Message}");
            }
        }


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

   

        public void ChangeDateFormatCaseExterior(string rutaArchivo, int columna, int nHoja, int filaIndexStart)
        {
            try
            {
                using (FileStream archivo = new FileStream(rutaArchivo, FileMode.Open))
                {
                    IWorkbook libro = new XSSFWorkbook(archivo);
                    string nombreHoja = libro.GetSheetAt(nHoja).SheetName;
                    ISheet hoja = libro.GetSheet(nombreHoja);

                    // Crear estilo de fecha
                    ICellStyle estiloFecha = libro.CreateCellStyle();
                    IDataFormat formatoDatos = libro.CreateDataFormat();
                    estiloFecha.DataFormat = formatoDatos.GetFormat("dd/MM/yyyy");

                    for (int filaIndex = filaIndexStart; filaIndex <= hoja.LastRowNum; filaIndex++)
                    {
                        IRow fila = hoja.GetRow(filaIndex);
                        if (fila != null)
                        {
                            ICell celda = fila.GetCell(columna - 1);
                            if (celda != null && celda.CellType == CellType.Numeric)
                            {
                                if (DateUtil.IsCellDateFormatted(celda))
                                {
                                    DateTime fecha = ObtenerValorCeldaFecha(celda);
                                    if (fecha.ToString("yy") != fecha.ToString("yyyy").Substring(2, 2))
                                    {
                                        celda.CellStyle = estiloFecha;
                                    }
                                }
                                else
                                {
                                    double valorNumerico = celda.NumericCellValue;
                                    DateTime fechaBase = new DateTime(1899, 12, 30); // Fecha base para Excel
                                    DateTime fecha = fechaBase.AddDays(valorNumerico);

                                    // Ajuste para el año 1900 (NPOI lo maneja incorrectamente)
                                    if (fecha.Year == 1900 && valorNumerico < 60)
                                    {
                                        fecha = fechaBase.AddDays(valorNumerico + 1);
                                    }

                                    if (fecha.ToString("yy") != fecha.ToString("yyyy").Substring(2, 2))
                                    {
                                        celda.SetCellValue(fecha);
                                        celda.CellStyle = estiloFecha;
                                    }
                                }
                            }
                        }
                    }

                    using (FileStream archivoSalida = new FileStream(rutaArchivo, FileMode.Create))
                    {
                        libro.Write(archivoSalida);
                    }
                }
                Console.WriteLine($"Columna {columna} formateada exitosamente.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error al formatear columna: {ex.Message}");
            }
        }
        public void ConvertColumnToGeneral(string rutaArchivo, int columnaReferencia, int nHoja)
        {
            try
            {
                using (FileStream archivo = new FileStream(rutaArchivo, FileMode.Open))
                {
                    IWorkbook libro = new XSSFWorkbook(archivo);
                    string nombreHoja = libro.GetSheetAt(nHoja).SheetName;
                    ISheet hoja = libro.GetSheet(nombreHoja);

                    for (int filaIndex = 0; filaIndex <= hoja.LastRowNum; filaIndex++)
                    {
                        IRow fila = hoja.GetRow(filaIndex);
                        if (fila != null)
                        {
                            ICell celdaReferencia = fila.GetCell(columnaReferencia - 1);
                            if (celdaReferencia != null && celdaReferencia.CellType == CellType.String)
                            {
                                string valorReferencia = celdaReferencia.StringCellValue;
                                if (double.TryParse(valorReferencia, out double numero))
                                {
                                    // Si se puede convertir a número, cambiar el tipo de celda a numérico
                                    celdaReferencia.SetCellValue(numero);
                                    celdaReferencia.SetCellType(CellType.Numeric);
                                }
                                else
                                {
                                    //Si no se puede convertir a número, dejarlo como string, para que excel lo trate como general.
                                    celdaReferencia.SetCellValue(valorReferencia);
                                    celdaReferencia.SetCellType(CellType.String);
                                }
                            }
                        }
                    }

                    using (FileStream archivoSalida = new FileStream(rutaArchivo, FileMode.Create))
                    {
                        libro.Write(archivoSalida);
                    }

                    Console.WriteLine($"Columna {columnaReferencia} convertida a formato general exitosamente.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error al convertir columna a formato general: {ex.Message}");
            }
        }

        public void HighlightRow(string rutaArchivo, int indiceFila, int nHoja)
        {
            try
            {
                using (FileStream archivo = new FileStream(rutaArchivo, FileMode.Open))
                {
                    IWorkbook libro = new XSSFWorkbook(archivo);
                    string nombreHoja = libro.GetSheetAt(nHoja).SheetName;
                    ISheet hoja = libro.GetSheet(nombreHoja);

                    ICellStyle estiloRojo = libro.CreateCellStyle();
                    estiloRojo.FillForegroundColor = IndexedColors.Red.Index;
                    estiloRojo.FillPattern = FillPattern.SolidForeground;

                    IRow fila = hoja.GetRow(indiceFila);
                    if (fila != null)
                    {
                        for (int columnaIndex = 0; columnaIndex <= 5; columnaIndex++)
                        {
                            ICell celda = fila.GetCell(columnaIndex);
                            if (celda == null)
                            {
                                celda = fila.CreateCell(columnaIndex);
                            }

                            // Copiar el estilo original de la celda
                            ICellStyle estiloOriginal = libro.CreateCellStyle();
                            estiloOriginal.CloneStyleFrom(celda.CellStyle);

                            // Aplicar el estilo de resaltado
                            celda.CellStyle = estiloRojo;

                            // Clonar el estilo rojo y conservar el estilo original.
                            ICellStyle estiloRojoClonado = libro.CreateCellStyle();
                            estiloRojoClonado.CloneStyleFrom(estiloRojo);

                            // combinamos el estilo rojo con el estilo original
                            estiloRojoClonado.CloneStyleFrom(estiloOriginal);
                            estiloRojoClonado.FillForegroundColor = IndexedColors.Red.Index;
                            estiloRojoClonado.FillPattern = FillPattern.SolidForeground;

                            //Aplicamos el estilo combinado a la celda.
                            celda.CellStyle = estiloRojoClonado;
                        }
                    }

                    using (FileStream archivoSalida = new FileStream(rutaArchivo, FileMode.Create))
                    {
                        libro.Write(archivoSalida);
                    }
                }
                Console.WriteLine($"Fila {indiceFila} resaltada exitosamente.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error al resaltar fila: {ex.Message}");
            }
        }

        public List<string> CopiarColumnaFechasComoStrings(string rutaArchivo, int sheetName, int columna)
        {
            List<string> listaStrings = new List<string>();

            try
            {
                using (FileStream archivo = new FileStream(rutaArchivo, FileMode.Open))
                {
                    IWorkbook libro = new XSSFWorkbook(archivo);
                    string nombreHoja = libro.GetSheetAt(sheetName).SheetName;
                    ISheet hoja = libro.GetSheet(nombreHoja);
                    int ultimaFila = hoja.LastRowNum;

                    for (int i = 0; i <= ultimaFila; i++)
                    {
                        IRow fila = hoja.GetRow(i);
                        if (fila != null)
                        {
                            ICell celda = fila.GetCell(columna);
                            if (celda == null || celda.Equals(""))
                            {
                                listaStrings.Add(""); // Agregar cadena vacía si la celda es nula
                            }
                            if (celda != null && !celda.Equals(""))
                            {
                                // Convertir el valor de la celda a string
                                string valorCelda = ObtenerValorCeldaComoString(celda);
                                listaStrings.Add(valorCelda);
                            }
  
                        }
                        else
                        {
                            listaStrings.Add(""); // Agregar cadena vacía si la fila es nula
                        }
                    }
                }

                return listaStrings;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error al copiar: {ex.Message}");
                return null;
            }
        }

        private string ObtenerValorCeldaComoString(ICell celda)
        {
            switch (celda.CellType)
            {
                case CellType.Boolean:
                    return celda.BooleanCellValue.ToString();
                case CellType.Numeric:
                 
                        return celda.NumericCellValue.ToString();
                case CellType.String:
                    return celda.StringCellValue;
                case CellType.Formula:
                    try
                    {
                        return celda.StringCellValue;
                    }
                    catch (Exception)
                    {
                        return celda.NumericCellValue.ToString();
                    }
                case CellType.Blank:
                    return "";
                case CellType.Error:
                    return celda.ErrorCellValue.ToString();
                default:
                    return "";
            }
        }

        public void ChangeCellTextFromListInReverseOrder(string rutaArchivo, int columna, int nHoja, List<string> listaColumna1)
        {
            try
            {
                using (FileStream archivo = new FileStream(rutaArchivo, FileMode.Open))
                {
                    IWorkbook libro = new XSSFWorkbook(archivo);
                    string nombreHoja = libro.GetSheetAt(nHoja).SheetName;
                    ISheet hoja = libro.GetSheet(nombreHoja);

                    int inverseIndex = listaColumna1.Count();

                    for (int i = 0; i < listaColumna1.Count; i++)
                    {
                        IRow filaObj = hoja.GetRow(i);
                        ICell celda = filaObj.GetCell(columna);

                        if (celda == null)
                        {
                            celda = filaObj.CreateCell(columna);
                        }

                        inverseIndex = inverseIndex - 1;

                        // 6. Cambiar el texto de la celda
                        celda.SetCellValue(listaColumna1[inverseIndex]);


                    }


 

                    using (FileStream archivoSalida = new FileStream(rutaArchivo, FileMode.Create))
                    {
                        libro.Write(archivoSalida);
                    }
                }

                Console.WriteLine("Texto de celda cambiado, formato y estilo aplicados exitosamente.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error al cambiar texto de celda o aplicar formato/estilo: " + ex.Message);
            }
        }

        public void ChangeCellTextFromListInTheSameOrder(string rutaArchivo, int columna, int nHoja, List<string> listaColumna1)
        {
            try
            {
                using (FileStream archivo = new FileStream(rutaArchivo, FileMode.Open))
                {
                    IWorkbook libro = new XSSFWorkbook(archivo);
                    string nombreHoja = libro.GetSheetAt(nHoja).SheetName;
                    ISheet hoja = libro.GetSheet(nombreHoja);


                    for (int i = 0; i < listaColumna1.Count; i++)
                    {
                        IRow filaObj = hoja.GetRow(i);
                        ICell celda = filaObj.GetCell(columna);

                        if (celda == null)
                        {
                            celda = filaObj.CreateCell(columna);
                        }

                        // 6. Cambiar el texto de la celda
                        celda.SetCellValue(listaColumna1[i]);


                    }




                    using (FileStream archivoSalida = new FileStream(rutaArchivo, FileMode.Create))
                    {
                        libro.Write(archivoSalida);
                    }
                }

                Console.WriteLine("Texto de celda cambiado, formato y estilo aplicados exitosamente.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error al cambiar texto de celda o aplicar formato/estilo: " + ex.Message);
            }
        }

        public void ReemplazarCeldasEnBlancoConCero(string rutaArchivo, int sheetName, int columnaIndex)
        {
            try
            {
                using (FileStream archivo = new FileStream(rutaArchivo, FileMode.Open))
                {
                    IWorkbook libro = new XSSFWorkbook(archivo);
                    ISheet hoja = libro.GetSheetAt(sheetName);
                    int ultimaFila = hoja.LastRowNum;

                    // Obtener el estilo de la celda (2, 1)
                    ICell celdaEstilo = hoja.GetRow(2)?.GetCell(1); // Usar operador condicional null
                    ICellStyle estilo = libro.CreateCellStyle();
                    IDataFormat formato = libro.CreateDataFormat();

                    if (celdaEstilo != null)
                    {
                        estilo = CopyCellStyle(celdaEstilo.CellStyle, libro);
                        estilo.DataFormat = formato.GetFormat("0;-0;;@");
                    }
                    else
                    {
                        // Crear un estilo predeterminado si celdaEstilo es nulo
                        estilo.DataFormat = formato.GetFormat("0;-0;;@");
                    }

                    for (int rowIndex = 0; rowIndex <= ultimaFila; rowIndex++)
                    {
                        IRow fila = hoja.GetRow(rowIndex);
                        if (fila != null)
                        {
                            ICell celda = fila.GetCell(columnaIndex);
                            if (celda == null || celda.CellType == CellType.Blank)
                            {
                                if (celda == null)
                                {
                                    celda = fila.CreateCell(columnaIndex);
                                }

                                celda.SetCellValue(0);
                                celda.CellStyle = estilo;
                            }
                            else if (celda.CellType == CellType.String && string.IsNullOrEmpty(celda.StringCellValue))
                            {
                                celda.SetCellValue(0);
                                celda.CellStyle = estilo;
                            }
                            else if (ObtenerValorCeldaString(celda).Equals("") || ObtenerValorCeldaComoString(celda) == "" || ObtenerValorCeldaComoString(celda).Equals(""))
                            {
                                celda.SetCellValue(0);
                                celda.CellStyle = estilo;
                            }
                            //else if (celda.NumericCellValue == 0)
                            //{
                            //    celda.CellStyle = estilo;
                            //}
                        }
                    }

                    using (FileStream archivoSalida = new FileStream(rutaArchivo, FileMode.Create))
                    {
                        libro.Write(archivoSalida);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public void ConvertirColumnaAString(string rutaArchivo, int sheetName, int columnaIndex)
        {
            try
            {
                using (FileStream archivo = new FileStream(rutaArchivo, FileMode.Open))
                {
                    IWorkbook libro = new XSSFWorkbook(archivo);
                    ISheet hoja = libro.GetSheetAt(sheetName);
                    int ultimaFila = hoja.LastRowNum;

                    for (int rowIndex = 0; rowIndex <= ultimaFila; rowIndex++)
                    {
                        IRow fila = hoja.GetRow(rowIndex);
                        if (fila != null)
                        {
                            ICell celda = fila.GetCell(columnaIndex);
                            if (celda != null)
                            {
                                string valorString = ObtenerValorCeldaComoString(celda);
                                celda.SetCellValue(valorString);
                            }
                        }
                    }

                    using (FileStream archivoSalida = new FileStream(rutaArchivo, FileMode.Create))
                    {
                        libro.Write(archivoSalida);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
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
                                    string valorFecha = ObtenerValorCeldaComoString(celdaFecha);
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



    }

}







    





using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Ajustador_de_formatos_Excel_de_movimientos_bancarios.BankFormatsFixerFunctions
{
    internal class Banesco
    {
                
        public int bankValidator(string rutaArchivo)
        {
            ExcelModifyFunctions modifyFunctions = new ExcelModifyFunctions();
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
                        string fecha = ExcelModifyFunctions.getValueCellString(fila.GetCell(0)).Trim().ToLower();
                        string fechaValidacion = ExcelModifyFunctions.getValueCellString(fila.GetCell(1)).Trim().ToLower();
                        string referencia = ExcelModifyFunctions.getValueCellString(fila.GetCell(1)).Trim().ToLower(); // Normalizar referencia
                        string descripcion = ExcelModifyFunctions.getValueCellString(fila.GetCell(2)).Trim().ToLower(); // Normalizar descripción
                        string monto = ExcelModifyFunctions.getValueCellString(fila.GetCell(3)).Trim().ToLower();
                        string balance = ExcelModifyFunctions.getValueCellString(fila.GetCell(4)).Trim().ToLower();

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

        public void fixFormat(TextBox ExcelFilePath)
        {
            ExcelModifyFunctions modifyFunctions = new ExcelModifyFunctions();

            // El archivo no está abierto por otro proceso se procede a la ejecución

            modifyFunctions.InsertColumnBetweenTwoCaseBanesco(ExcelFilePath.Text, 5);
            modifyFunctions.InsertColumnBetweenTwoVersionC2(ExcelFilePath.Text, 2);

            //Ajustando columna "Fecha de validación"
            modifyFunctions.AdjustColumnWidth(ExcelFilePath.Text, 2, 20, 0);
            //Ajustando columna "Descripción"
            modifyFunctions.AdjustColumnWidth(ExcelFilePath.Text, 4, 50, 0);
            //Ajustando columna "Referencia"
            modifyFunctions.AdjustColumnWidth(ExcelFilePath.Text, 3, 14, 0);

            modifyFunctions.MoveNegativesNumbersCaseBanesco(ExcelFilePath.Text, 5, 6, 0);

            //Modificación personal para facilitar la ubicación:
            /* 
              * Las columnas se cuentan desde 0, las filas no.
              * No modifiqué el conteo en las columnas, porque la mayor parte
              * de las modificaciones a realizar no involucran un conteo demasiado
              * variado, en el caso de las filas me estaba complicando en las comparaciones.
              * 
              * */

            modifyFunctions.ChangeCellTextWithFormatAndStyle(ExcelFilePath.Text, 0, 1, "Fecha de validación");
            modifyFunctions.ChangeCellTextWithFormatAndStyle(ExcelFilePath.Text, 0, 4, "Ingresos");
            modifyFunctions.ChangeCellTextWithFormatAndStyle(ExcelFilePath.Text, 0, 5, "Egresos");
            modifyFunctions.ChangeCellTextWithFormatAndStyle(ExcelFilePath.Text, 0, 16, "Archivo modificado");
            modifyFunctions.ChangeCellTextWithFormatAndStyle(ExcelFilePath.Text, 0, 17, "Fecha de modificación:" + DateTime.Now.ToString());

            //Reparando formato de las celdas en blanco (para que no se dañe la fórmula)

            modifyFunctions.replaceEmptyCellsWithZero(ExcelFilePath.Text, 0, 4);
            modifyFunctions.replaceEmptyCellsWithZero(ExcelFilePath.Text, 0, 5);
            modifyFunctions.replaceEmptyCellsWithZero(ExcelFilePath.Text, 0, 6);

            MessageBox.Show("Ajustes realizados exitosamente", "Proceso finalizado", MessageBoxButtons.OK, MessageBoxIcon.Information);

        }

    }
}

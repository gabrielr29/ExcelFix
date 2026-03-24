using Ajustador_de_formatos_Excel_de_movimientos_bancarios.BusinessLogic;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Ajustador_de_formatos_Excel_de_movimientos_bancarios.BankFormatsFixerFunctions
{
    internal class BDV
    {
        

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
                        string fecha = ExcelModifyFunctions.getValueCellString(fila.GetCell(0));
                        string fechaValidacion = ExcelModifyFunctions.getValueCellString(fila.GetCell(1)).Trim().ToLower();
                        string referencia = ExcelModifyFunctions.getValueCellString(fila.GetCell(1)).Trim().ToLower(); // Normalizar referencia
                        string concepto = ExcelModifyFunctions.getValueCellString(fila.GetCell(2)).Trim().ToLower(); // Normalizar descripción
                        string saldo = ExcelModifyFunctions.getValueCellString(fila.GetCell(3)).Trim().ToLower();
                        string monto = ExcelModifyFunctions.getValueCellString(fila.GetCell(4)).Trim().ToLower();
                        string tipoMov = ExcelModifyFunctions.getValueCellString(fila.GetCell(5)).Trim().ToLower();
                        string rif = ExcelModifyFunctions.getValueCellString(fila.GetCell(6)).Trim().ToLower();
                        string numeroCuenta = ExcelModifyFunctions.getValueCellString(fila.GetCell(7)).Trim().ToLower();


                        // Crear una clave única para la fila
                        string filaHash = $"{fecha}{referencia}{concepto}{saldo}{monto}{tipoMov}{rif}{numeroCuenta}";

                        if (filaHash.Equals("fechareferenciaconceptosaldomontotipomovimientorifnumerocuenta"))
                        {
                            // CASO 1, EL DOCUMENTO SI ES UN FORMATO DE CONSULTA DE MOVIMIENTOS BCO VNZLA, SIN MODIFICAR
                            return 1;
                        }
                        else if (fechaValidacion.Equals("fecha de validación"))
                        {
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

        public void fixFormat(TextBox ExcelFilePath)
        {
            ExcelModifyFunctions modifyFunctions = new ExcelModifyFunctions();

            modifyFunctions.InsertColumnBetweenTwoCaseBanesco(ExcelFilePath.Text, 5);
            modifyFunctions.MoveColumnsCaseBVnzlaBExterior(ExcelFilePath.Text, 6, 5, 3, 0);
            modifyFunctions.MoveColumnsCaseBVnzlaBExterior(ExcelFilePath.Text, 4, 6, 3, 0);
            modifyFunctions.DeleteColumnAndMove(ExcelFilePath.Text, 4);
            modifyFunctions.DeleteColumnAndMove(ExcelFilePath.Text, 6);

            modifyFunctions.InsertColumnBetweenTwoVersionC2(ExcelFilePath.Text, 2);
            modifyFunctions.ChangeCellTextWithFormatAndStyle(ExcelFilePath.Text, 0, 1, "Fecha de validación");
            modifyFunctions.InsertColumnBetweenTwoVersionC2(ExcelFilePath.Text, 6);

            modifyFunctions.MoveColumnsCaseBVnzlaBExterior(ExcelFilePath.Text, 5, 6, 1, 0);

            //Ajustando columna "Fecha de validación"
            modifyFunctions.AdjustColumnWidth(ExcelFilePath.Text, 2, 20, 0);
            //Ajustando columna "Referencias"
            modifyFunctions.AdjustColumnWidth(ExcelFilePath.Text, 3, 20, 0);
            //Ajustando columna "Concepto"
            modifyFunctions.AdjustColumnWidth(ExcelFilePath.Text, 4, 50, 0);

            //Reajustando fuente restantes para mejorar la estética
            modifyFunctions.ChangeCellTextWithFormatAndStyle(ExcelFilePath.Text, 0, 0, "Fecha");
            modifyFunctions.ChangeCellTextWithFormatAndStyle(ExcelFilePath.Text, 0, 2, "Referencia");
            modifyFunctions.ChangeCellTextWithFormatAndStyle(ExcelFilePath.Text, 0, 3, "Concepto");


            //Borrando las columnas con problemas de formato del banco (no cambian correctamente de General - Número)
            //y moviendo los datos

            modifyFunctions.InsertColumnBetweenTwoVersionC2(ExcelFilePath.Text, 7);
            modifyFunctions.InsertColumnBetweenTwoVersionC2(ExcelFilePath.Text, 6);
            modifyFunctions.MoveColumnsCaseBVnzlaBExterior(ExcelFilePath.Text, 5, 6, 2, 0);
            modifyFunctions.MoveColumnsCaseBVnzlaBExterior(ExcelFilePath.Text, 9, 8, 3, 0);


            //Ajustando columna "Saldo"
            modifyFunctions.AdjustColumnWidth(ExcelFilePath.Text, 7, 15, 0);

            //Eliminando columnas problemáticas
            modifyFunctions.DeleteColumnAndMove(ExcelFilePath.Text, 5);
            modifyFunctions.CleanColumn(ExcelFilePath.Text, 8, 0);
            modifyFunctions.CleanColumn(ExcelFilePath.Text, 9, 0);
            modifyFunctions.CleanColumn(ExcelFilePath.Text, 10, 0);

            //Ajustando las fuentes y mejorando la estética
            modifyFunctions.ChangeCellTextWithFormatAndStyle(ExcelFilePath.Text, 0, 4, "Ingresos");
            modifyFunctions.ChangeCellTextWithFormatAndStyle(ExcelFilePath.Text, 0, 5, "Egresos");
            modifyFunctions.ChangeCellTextWithFormatAndStyle(ExcelFilePath.Text, 0, 6, "Saldo");

            //Reparando formato de las celdas en blanco (para que no se dañe la fórmula)

            modifyFunctions.replaceEmptyCellsWithZero(ExcelFilePath.Text, 0, 4);
            modifyFunctions.replaceEmptyCellsWithZero(ExcelFilePath.Text, 0, 5);
            modifyFunctions.replaceEmptyCellsWithZero(ExcelFilePath.Text, 0, 6);

            MessageBox.Show("Ajustes realizados exitosamente", "Proceso finalizado", MessageBoxButtons.OK, MessageBoxIcon.Information);


            /*                          
             * La diferencia entre InsertColumnBetweenTwoCaseBanesco(ExcelFilePath.Text, 5); y 
             * InsertColumnBetweenTwoCaseBanescoC2(ExcelFilePath.Text, 2); es que la primera inserta
             * copiando todos los datos de la columna a la izquierda, facilitando el traslado de los datos
             * mientras que la segunda simplemente inserta una columna en blanco.   
             *                               
             * */



        }




    }
}

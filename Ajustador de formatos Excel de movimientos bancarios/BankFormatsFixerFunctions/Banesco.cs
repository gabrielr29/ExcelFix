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




    }
}

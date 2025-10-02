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




    }
}

using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Ajustador_de_formatos_Excel_de_movimientos_bancarios.BankFormatsFixerFunctions
{
    internal class Exterior
    {

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



    }
}

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

        ExcelModifyFunctions modifyFunctions = new ExcelModifyFunctions();

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










    }
}

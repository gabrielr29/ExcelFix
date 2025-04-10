using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Ajustador_de_formatos_Excel_de_movimientos_bancarios
{
    internal class MovValidatorFunctions
    {

        public static void UploadAndFilterExcel(string rutaArchivoSeleccionado)
        {

        }

        private static string FormatValidationDate(DateTime validationDate)
        {
            if (!validationDate.Equals(DateTime.MinValue))
            {
                string fechaValidacionFormateada = validationDate.ToString("dd/MM/yyyy");
                return fechaValidacionFormateada;
            }
            return "";
        }

        public static List<string> SearchByReferenceAndMount(string rutaArchivo, int startedRowToRevision, string referenciaBusqueda, decimal montoBusqueda)
        {
            List<string> datosFilaEncontrada = new List<string>();
            ExcelModifyFunctions functions = new ExcelModifyFunctions();

            try
            {
                using (FileStream archivo = new FileStream(rutaArchivo, FileMode.Open, FileAccess.Read))
                {
                    IWorkbook libro = new XSSFWorkbook(archivo);
                    ISheet hoja = libro.GetSheetAt(0);

                    for (int i = startedRowToRevision; i <= hoja.LastRowNum; i++)
                    {
                        IRow fila = hoja.GetRow(i);
                        if (fila != null)
                        {
                            string referencia = functions.ObtenerValorCeldaString(fila.GetCell(2)).Trim().ToLower();
                            decimal ingresos = functions.ObtenerValorCeldaDecimal(fila.GetCell(4));
                            decimal egresos = functions.ObtenerValorCeldaDecimal(fila.GetCell(5));

                            if (referencia == referenciaBusqueda.Trim().ToLower() && (ingresos == montoBusqueda || egresos == montoBusqueda))
                            {
                                DateTime fecha = functions.ObtenerValorCeldaFecha(fila.GetCell(0));
                                string fechaFormateada = fecha.ToString("dd/MM/yyyy");
                                DateTime fechaValidacion = functions.ObtenerValorCeldaFecha(fila.GetCell(1));
                                string fechaValidacionFormateada = FormatValidationDate(fechaValidacion);
                                string descripcion = functions.ObtenerValorCeldaString(fila.GetCell(3)).Trim();
                                string numeroFactura = functions.ObtenerValorCeldaString(fila.GetCell(7));
                                string codigoCliente = functions.ObtenerValorCeldaString(fila.GetCell(8));

                                datosFilaEncontrada.Add($"{fechaFormateada}");
                                datosFilaEncontrada.Add($"{fechaValidacionFormateada}");
                                datosFilaEncontrada.Add($"{referencia}");
                                datosFilaEncontrada.Add($"{descripcion}");
                                datosFilaEncontrada.Add($"{ingresos}");
                                datosFilaEncontrada.Add($"{egresos}");
                                datosFilaEncontrada.Add($"{numeroFactura}");
                                datosFilaEncontrada.Add($"{codigoCliente}");
                                datosFilaEncontrada.Add($"{i}");

                                break; // Solo buscamos la primera coincidencia
                            }
                        }
                    }

                    if (datosFilaEncontrada.Count == 0)
                    {
                        datosFilaEncontrada.Add("No se encontró ninguna fila con la referencia y monto indicados.");
                    }
                }
            }
            catch (Exception ex)
            {
                datosFilaEncontrada.Add("Error al buscar la fila: " + ex.Message);
            }

            return datosFilaEncontrada;
        }

        public static List<string> SearchByMount(string rutaArchivo, int startedRowToRevision, decimal montoBusqueda)
        {
            List<string> resultados = new List<string>();
            ExcelModifyFunctions functions = new ExcelModifyFunctions();

            try
            {
                using (FileStream archivo = new FileStream(rutaArchivo, FileMode.Open, FileAccess.Read))
                {
                    IWorkbook libro = new XSSFWorkbook(archivo);
                    ISheet hoja = libro.GetSheetAt(0);

                    for (int i = startedRowToRevision; i <= hoja.LastRowNum; i++)
                    {
                        IRow fila = hoja.GetRow(i);
                        if (fila != null)
                        {
                            decimal ingresos = functions.ObtenerValorCeldaDecimal(fila.GetCell(4));
                            decimal egresos = functions.ObtenerValorCeldaDecimal(fila.GetCell(5));

                            if (ingresos == montoBusqueda || egresos == montoBusqueda)
                            {
                                DateTime fecha = functions.ObtenerValorCeldaFecha(fila.GetCell(0));
                                string fechaFormateada = fecha.ToString("dd/MM/yyyy");
                                DateTime fechaValidacion = functions.ObtenerValorCeldaFecha(fila.GetCell(1));
                                string fechaValidacionFormateada = FormatValidationDate(fechaValidacion);
                                string referencia = functions.ObtenerValorCeldaString(fila.GetCell(2)).Trim();
                                string descripcion = functions.ObtenerValorCeldaString(fila.GetCell(3)).Trim();
                                string numeroFactura = functions.ObtenerValorCeldaString(fila.GetCell(7));
                                string codigoCliente = functions.ObtenerValorCeldaString(fila.GetCell(8));

                                resultados.Add($"{fechaFormateada}");
                                resultados.Add($"{fechaValidacionFormateada}");
                                resultados.Add($"{referencia}");
                                resultados.Add($"{descripcion}");
                                resultados.Add($"{ingresos}");
                                resultados.Add($"{egresos}");
                                resultados.Add($"{numeroFactura}");
                                resultados.Add($"{codigoCliente}");
                                resultados.Add($"{i}");
                                
                            }
                        }
                    }

                    if (resultados.Count == 0)
                        resultados.Add("No se encontraron coincidencias con el monto indicado.");
                }
            }
            catch (Exception ex)
            {
                resultados.Add("Error al buscar por monto: " + ex.Message);
            }

            return resultados;
        }

        public static List<string> SearchByReference(string rutaArchivo, int startedRowToRevision, string referenciaBusqueda)
        {
            List<string> resultados = new List<string>();
            ExcelModifyFunctions functions = new ExcelModifyFunctions();

            try
            {
                using (FileStream archivo = new FileStream(rutaArchivo, FileMode.Open, FileAccess.Read))
                {
                    IWorkbook libro = new XSSFWorkbook(archivo);
                    ISheet hoja = libro.GetSheetAt(0);

                    for (int i = startedRowToRevision; i <= hoja.LastRowNum; i++)
                    {
                        IRow fila = hoja.GetRow(i);
                        if (fila != null)
                        {
                            string referencia = functions.ObtenerValorCeldaString(fila.GetCell(2)).Trim().ToLower();

                            if (referencia == referenciaBusqueda.Trim().ToLower())
                            {
                                DateTime fecha = functions.ObtenerValorCeldaFecha(fila.GetCell(0));
                                string fechaFormateada = fecha.ToString("dd/MM/yyyy");
                                DateTime fechaValidacion = functions.ObtenerValorCeldaFecha(fila.GetCell(1));
                                string fechaValidacionFormateada = FormatValidationDate(fechaValidacion);
                                string descripcion = functions.ObtenerValorCeldaString(fila.GetCell(3)).Trim();
                                decimal ingresos = functions.ObtenerValorCeldaDecimal(fila.GetCell(4));
                                decimal egresos = functions.ObtenerValorCeldaDecimal(fila.GetCell(5));
                                string numeroFactura = functions.ObtenerValorCeldaString(fila.GetCell(7));
                                string codigoCliente = functions.ObtenerValorCeldaString(fila.GetCell(8));

                                resultados.Add($"{fechaFormateada}");
                                resultados.Add($"{fechaValidacionFormateada}");
                                resultados.Add($"{referencia}");
                                resultados.Add($"{descripcion}");
                                resultados.Add($"{ingresos}");
                                resultados.Add($"{egresos}");
                                resultados.Add($"{numeroFactura}");
                                resultados.Add($"{codigoCliente}");
                                resultados.Add($"{i}");
                                
                            }
                        }
                    }

                    if (resultados.Count == 0)
                        resultados.Add("No se encontraron coincidencias con la referencia indicada.");
                }
            }
            catch (Exception ex)
            {
                resultados.Add("Error al buscar por referencia: " + ex.Message);
            }

            return resultados;
        }


        public static void ReplaceDataGridViewValues(DataGridView dataGridView1, List<string> myList)
        {
            // Verifica que la lista tenga suficientes elementos para reemplazar todas las celdas
            if (myList.Count >= dataGridView1.RowCount * dataGridView1.ColumnCount)
            {
                int listIndex = 0;

                // Itera a través de las filas y columnas del DataGridView
                for (int rowIndex = 0; rowIndex < dataGridView1.RowCount; rowIndex++)
                {
                    for (int columnIndex = 0; columnIndex < dataGridView1.ColumnCount; columnIndex++)
                    {
                        // Reemplaza el valor de la celda con el valor correspondiente de la lista
                        dataGridView1.Rows[rowIndex].Cells[columnIndex].Value = myList[listIndex];
                        listIndex++;
                    }
                }
            }
            else
            {
                MessageBox.Show($"La lista no tiene suficientes elementos para reemplazar todos los valores del DataGridView. {myList[0]}");
            }
        }

        public static void ReplaceDataGridViewValues2(DataGridView dataGridView1, List<string> myList)
        {
            // Validación: Asegurar que la lista tenga una cantidad de elementos múltiplo de 9
            if (myList.Count % 9 != 0)
            {
                MessageBox.Show($"{myList[0]}");
                return;
            }

            int totalFilas = myList.Count / 9;

            // Limpia cualquier fila anterior
            dataGridView1.Rows.Clear();

            // Agrega y llena las filas según los datos
            for (int i = 0; i < totalFilas; i++)
            {
                dataGridView1.Rows.Add(); // Agrega nueva fila

                for (int j = 0; j < 9 && j < dataGridView1.ColumnCount; j++)
                {
                    int dataIndex = i * 9 + j; // Este es el índice correcto para cada celda
                    dataGridView1.Rows[i].Cells[j].Value = myList[dataIndex];
                }
            }

            
        }

        public static void UpdateCellsByRow(string rutaArchivo, int numeroFila, DateTime fechaValidacion, string numeroFactura, string codigoCliente)
        {
            try
            {
                using (FileStream archivo = new FileStream(rutaArchivo, FileMode.Open, FileAccess.ReadWrite))
                {
                    IWorkbook libro = new XSSFWorkbook(archivo);
                    ISheet hoja = libro.GetSheetAt(0);

                    IDataFormat formato = libro.CreateDataFormat();

                    // Estilo para formato de fecha dd/MM/yyyy
                    ICellStyle estiloFecha = libro.CreateCellStyle();
                    
                    estiloFecha.FillPattern = FillPattern.SolidForeground;
                    estiloFecha.FillForegroundColor = IndexedColors.LightBlue.Index; // Color azul
                    estiloFecha.DataFormat = formato.GetFormat("dd/MM/yyyy");

                    //Estilo con formato general y coloreado Azul

                    IRow fila = hoja.GetRow(numeroFila);
                    if (fila == null)
                        fila = hoja.CreateRow(numeroFila);

                    // Fecha de Validación (columna 1)
                    ICell celdaFechaValidacion = fila.GetCell(1) ?? fila.CreateCell(1);
                    celdaFechaValidacion.SetCellValue(fechaValidacion.Date);
                    estiloFecha.BorderLeft = celdaFechaValidacion.CellStyle.BorderLeft;
                    celdaFechaValidacion.CellStyle = estiloFecha;

                    // Número de Factura (columna 7)
                    ICell celdaFactura = fila.GetCell(7) ?? fila.CreateCell(7);
                    celdaFactura.SetCellValue(numeroFactura);
                    celdaFactura.CellStyle = CloneSyleAndFormat(libro, celdaFactura.CellStyle, formato, true);

                    // Código de Cliente (columna 8)
                    ICell celdaCliente = fila.GetCell(8) ?? fila.CreateCell(8);
                    celdaCliente.SetCellValue(codigoCliente);
                    celdaCliente.CellStyle = CloneSyleAndFormat(libro, celdaCliente.CellStyle, formato, true);

                    //Celdas restantes(columna 6, 5, 4, 3, 2, 0)
                    ICell celdaFechaOriginal = fila.GetCell(0) ?? fila.CreateCell(0);
                    celdaFechaOriginal.CellStyle = celdaFechaValidacion.CellStyle;
                    
                    ICell celdaReferencia = fila.GetCell(2) ?? fila.CreateCell(2);
                    celdaReferencia.CellStyle = CloneSyleAndFormat(libro, celdaReferencia.CellStyle, formato, false);

                    ICell celdaDescripcion = fila.GetCell(3) ?? fila.CreateCell(3);
                    celdaDescripcion.CellStyle = CloneSyleAndFormat(libro, celdaDescripcion.CellStyle, formato, false);

                    ICell celdaIngresos = fila.GetCell(4) ?? fila.CreateCell(4);
                    celdaIngresos.CellStyle = CloneSyleAndFormat(libro, celdaIngresos.CellStyle, formato, false);

                    ICell celdaEgresos = fila.GetCell(5) ?? fila.CreateCell(5);

                    celdaEgresos.CellStyle = CloneSyleAndFormat(libro, celdaEgresos.CellStyle, formato, false);

                    ICell celdaSaldo = fila.GetCell(6) ?? fila.CreateCell(6);

                    celdaSaldo.CellStyle = CloneSyleAndFormat(libro, celdaSaldo.CellStyle, formato, false);



                    // Guardar cambios
                    using (FileStream salida = new FileStream(rutaArchivo, FileMode.Create, FileAccess.Write))
                    {
                        libro.Write(salida);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al actualizar las celdas: " + ex.Message);
            }
        }

        // Función para clonar el estilo y aplicar el formato "General", manteniendo bordes y otros atributos pero añadiendo color azul
        private static ICellStyle CloneSyleAndFormat(IWorkbook libro, ICellStyle estiloOriginal, IDataFormat formato, bool generalOrOriginal)
        {
            ICellStyle nuevoEstilo = libro.CreateCellStyle();

            if (estiloOriginal != null)
            {
                // Copia los colores de borde
                nuevoEstilo.BottomBorderColor = estiloOriginal.BottomBorderColor;
                nuevoEstilo.TopBorderColor = estiloOriginal.TopBorderColor;
                nuevoEstilo.LeftBorderColor = estiloOriginal.LeftBorderColor;
                nuevoEstilo.RightBorderColor = estiloOriginal.RightBorderColor;

                // Copia alineación, fuente, etc.
                nuevoEstilo.Alignment = estiloOriginal.Alignment;
                nuevoEstilo.VerticalAlignment = estiloOriginal.VerticalAlignment;
                nuevoEstilo.WrapText = estiloOriginal.WrapText;
                nuevoEstilo.FillBackgroundColor = estiloOriginal.FillBackgroundColor;
                nuevoEstilo.ShrinkToFit = estiloOriginal.ShrinkToFit;
                nuevoEstilo.Indention = estiloOriginal.Indention;
                nuevoEstilo.Rotation = estiloOriginal.Rotation;

                // Sobreescribiendo sombreado al color necesario
                nuevoEstilo.FillPattern = FillPattern.SolidForeground;
                nuevoEstilo.FillForegroundColor = IndexedColors.LightBlue.Index;

                // Copia los bordes
                nuevoEstilo.BorderBottom = estiloOriginal.BorderBottom;
                nuevoEstilo.BorderTop = estiloOriginal.BorderTop;
                nuevoEstilo.BorderLeft = estiloOriginal.BorderLeft;
                nuevoEstilo.BorderRight = estiloOriginal.BorderRight;
            }

            // Asigna el formato "General" o el formato original
            nuevoEstilo.DataFormat = generalOrOriginal ? formato.GetFormat("General") : estiloOriginal.DataFormat;

            return nuevoEstilo;
        }


    }
}

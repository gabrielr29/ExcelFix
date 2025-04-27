using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;


namespace Ajustador_de_formatos_Excel_de_movimientos_bancarios
{
    internal class MovValidatorFunctions
    {

        private static string FormatValidationDate(DateTime validationDate)
        {
            if (!validationDate.Equals(DateTime.MinValue))
            {
                string fechaValidacionFormateada = validationDate.ToString("dd/MM/yyyy");
                return fechaValidacionFormateada;
            }
            
            return "";
        }

        public static DateTime ConvertirStringADateTime(string fechaString)
        {
            string formato = "dd/MM/yyyy";

            try
            {
                DateTime fechaDateTime = DateTime.ParseExact(fechaString, formato, CultureInfo.InvariantCulture, DateTimeStyles.None);
                return fechaDateTime;
            }
            catch (FormatException ex)
            {
                Console.WriteLine($"Error al convertir la fecha: {ex.Message}");
                return DateTime.MinValue; // Retorna DateTime.MinValue en caso de error.
            }
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

                            if (referencia == referenciaBusqueda.Trim().ToLower() && (ingresos == montoBusqueda && egresos == 0))
                            {

                                //DateTime fechaValidacion = functions.ObtenerValorCeldaFecha(fila.GetCell(1));
                                //string fechaValidacionFormateada = FormatValidationDate(fechaValidacion);

                                string descripcion = functions.ObtenerValorCeldaString(fila.GetCell(3)).Trim();
                                string numeroFactura = functions.ObtenerValorCeldaString(fila.GetCell(7));
                                string codigoCliente = functions.ObtenerValorCeldaString(fila.GetCell(8));

                                string fechaValidacionFormateada = "";
                                ICell fechaValidacionCell = fila.GetCell(1);

                                if (CheckCellType(fechaValidacionCell).Equals("Fecha"))
                                {
                                    DateTime fechaValidacion = functions.ObtenerValorCeldaFecha(fechaValidacionCell);
                                    fechaValidacionFormateada = fechaValidacion.ToString("dd/MM/yyyy");
                                }
                                else
                                {
                                    fechaValidacionFormateada = functions.ObtenerValorCeldaString(fechaValidacionCell);
                                }

                                if (CheckCellType(fila.GetCell(0)).Equals("Fecha"))
                                {
                                    DateTime fecha = functions.ObtenerValorCeldaFecha(fila.GetCell(0));
                                    string fechaFormateada = fecha.ToString("dd/MM/yyyy");
                                    datosFilaEncontrada.Add($"{fechaFormateada}");
                                }

                                else
                                {
                                    string fecha = functions.ObtenerValorCeldaString(fila.GetCell(0));
                                    datosFilaEncontrada.Add($"{fecha}");
                                }
                                                                
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

                            if (ingresos == montoBusqueda && egresos == 0)
                            {
                                
                                //DateTime fechaValidacion = functions.ObtenerValorCeldaFecha(fila.GetCell(1));
                                //string fechaValidacionFormateada = FormatValidationDate(fechaValidacion);

                                string fechaValidacionFormateada = "";
                                ICell fechaValidacionCell = fila.GetCell(1);

                                if (CheckCellType(fechaValidacionCell).Equals("Fecha"))
                                {
                                    DateTime fechaValidacion = functions.ObtenerValorCeldaFecha(fechaValidacionCell);
                                    fechaValidacionFormateada = fechaValidacion.ToString("dd/MM/yyyy");
                                }
                                else
                                {
                                    fechaValidacionFormateada = functions.ObtenerValorCeldaString(fechaValidacionCell);
                                }

                                string referencia = functions.ObtenerValorCeldaString(fila.GetCell(2)).Trim();
                                string descripcion = functions.ObtenerValorCeldaString(fila.GetCell(3)).Trim();
                                string numeroFactura = functions.ObtenerValorCeldaString(fila.GetCell(7));
                                string codigoCliente = functions.ObtenerValorCeldaString(fila.GetCell(8));

                                if (CheckCellType(fila.GetCell(0)).Equals("Fecha"))
                                {
                                    DateTime fecha = functions.ObtenerValorCeldaFecha(fila.GetCell(0));
                                    string fechaFormateada = fecha.ToString("dd/MM/yyyy");
                                    resultados.Add($"{fechaFormateada}");
                                }

                                else
                                {
                                    string fecha = functions.ObtenerValorCeldaString(fila.GetCell(0));
                                    resultados.Add($"{fecha}");
                                }
                                                                
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
                            decimal egresosToCompare = functions.ObtenerValorCeldaDecimal(fila.GetCell(5));

                            if (referencia == referenciaBusqueda.Trim().ToLower() && egresosToCompare == 0)
                            {
                                //DateTime fechaValidacion = functions.ObtenerValorCeldaFecha(fila.GetCell(1));
                                //string fechaValidacionFormateada = FormatValidationDate(fechaValidacion);

                                string descripcion = functions.ObtenerValorCeldaString(fila.GetCell(3)).Trim();
                                decimal ingresos = functions.ObtenerValorCeldaDecimal(fila.GetCell(4));
                                decimal egresos = functions.ObtenerValorCeldaDecimal(fila.GetCell(5));
                                string numeroFactura = functions.ObtenerValorCeldaString(fila.GetCell(7));
                                string codigoCliente = functions.ObtenerValorCeldaString(fila.GetCell(8));

                                string fechaValidacionFormateada = "";
                                ICell fechaValidacionCell = fila.GetCell(1);

                                if (CheckCellType(fechaValidacionCell).Equals("Fecha"))
                                {
                                    DateTime fechaValidacion = functions.ObtenerValorCeldaFecha(fechaValidacionCell);
                                    fechaValidacionFormateada = fechaValidacion.ToString("dd/MM/yyyy");
                                }
                                else
                                {
                                    fechaValidacionFormateada = functions.ObtenerValorCeldaString(fechaValidacionCell);
                                }

                                if (CheckCellType(fila.GetCell(0)).Equals("Fecha"))
                                {
                                    DateTime fecha = functions.ObtenerValorCeldaFecha(fila.GetCell(0));
                                    string fechaFormateada = fecha.ToString("dd/MM/yyyy");
                                    resultados.Add($"{fechaFormateada}");
                                }

                                else
                                {
                                    string fecha = functions.ObtenerValorCeldaString(fila.GetCell(0));
                                    resultados.Add($"{fecha}");
                                }

                                
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

        /*
         * Función actualizada para que considere las coincidencias a partir de los últimos 4 dígitos
         * o bien los últimos 6, agregando más flexibilidad.
         */

        public static List<string> SearchByReferenceIII(string rutaArchivo, int startedRowToRevision, string referenciaBusqueda)
        {
            List<string> resultados = new List<string>();

            string digitosBusqueda = referenciaBusqueda.Trim().ToLower();

            resultados = forLoopSearchByReferenceIII(startedRowToRevision, rutaArchivo, digitosBusqueda);

            return resultados;
        }


        public static List<string> SearchByReferenceandMountII(string rutaArchivo, int startedRowToRevision, string referenciaBusqueda, decimal montoBusqueda)
        {
            List<string> resultados = new List<string>();

            string digitosBusqueda = referenciaBusqueda.Trim().ToLower();

            resultados = forLoopSearchByReferenceandMountII(startedRowToRevision, rutaArchivo, digitosBusqueda, montoBusqueda);

            return resultados;
        }



        public static List<string> forLoopSearchByReferenceIII(int startedRowToRevision, string rutaArchivo, string digitosBusqueda)
        {
            List<string> resultados = new List<string>();
            ExcelModifyFunctions functions = new ExcelModifyFunctions();

            using (FileStream archivo = new FileStream(rutaArchivo, FileMode.Open, FileAccess.Read))
            {
                try
                {
                    IWorkbook libro = new XSSFWorkbook(archivo);
                    ISheet hoja = libro.GetSheetAt(0);

                    for (int i = startedRowToRevision; i <= hoja.LastRowNum; i++)
                    {
                        IRow fila = hoja.GetRow(i);
                        if (fila != null)
                        {
                            string referenciaCelda = functions.ObtenerValorCeldaString(fila.GetCell(2)).Trim().ToLower();
                            decimal egresosToCompare = functions.ObtenerValorCeldaDecimal(fila.GetCell(5));
                            ICell fechaCell = fila.GetCell(0);
                            ICell fechaValidacionCell = fila.GetCell(1);
                            string descripcion = functions.ObtenerValorCeldaString(fila.GetCell(3)).Trim();
                            decimal ingresos = functions.ObtenerValorCeldaDecimal(fila.GetCell(4));
                            string numeroFactura = functions.ObtenerValorCeldaString(fila.GetCell(7));
                            string codigoCliente = functions.ObtenerValorCeldaString(fila.GetCell(8));

                            bool coincidenciaUltimosDigitos = false;
                            if (referenciaCelda.EndsWith(digitosBusqueda))
                            {
                                coincidenciaUltimosDigitos = true;
                            }

                            if (coincidenciaUltimosDigitos && egresosToCompare == 0)
                            {

                                //DateTime fechaValidacion = functions.ObtenerValorCeldaFecha(fechaValidacionCell);
                                //string fechaValidacionFormateada = FormatValidationDate(fechaValidacion);

                                string fechaValidacionFormateada = "";

                                if (CheckCellType(fechaValidacionCell).Equals("Fecha"))
                                {
                                    DateTime fechaValidacion = functions.ObtenerValorCeldaFecha(fechaValidacionCell);
                                    fechaValidacionFormateada = fechaValidacion.ToString("dd/MM/yyyy");
                                }
                                else
                                {
                                    fechaValidacionFormateada = functions.ObtenerValorCeldaString(fechaValidacionCell);
                                }

                                string fechaFormateada = "";

                                if (CheckCellType(fechaCell).Equals("Fecha"))
                                {
                                    DateTime fecha = functions.ObtenerValorCeldaFecha(fechaCell);
                                    fechaFormateada = fecha.ToString("dd/MM/yyyy");
                                }
                                else
                                {
                                    fechaFormateada = functions.ObtenerValorCeldaString(fechaCell);
                                }


                                resultados.Add($"{fechaFormateada}");
                                resultados.Add($"{fechaValidacionFormateada}");
                                resultados.Add($"{referenciaCelda}");
                                resultados.Add($"{descripcion}");
                                resultados.Add($"{ingresos}");
                                resultados.Add($"{egresosToCompare}");
                                resultados.Add($"{numeroFactura}");
                                resultados.Add($"{codigoCliente}");
                                resultados.Add($"{i}");
                            }
                        }
                    }

                    if (resultados.Count == 0)
                    {
                        resultados.Add("No se encontraron coincidencias con la referencia indicada.");
                    }
                }
                catch (Exception ex)
                {
                    resultados.Add("Error al buscar por referencia: " + ex.Message);
                }


                return resultados;
            }
        }

        public static List<string> forLoopSearchByReferenceandMountII(int startedRowToRevision, string rutaArchivo, string digitosBusqueda, decimal montoBusqueda)
        {
            List<string> resultados = new List<string>();
            ExcelModifyFunctions functions = new ExcelModifyFunctions();

            using (FileStream archivo = new FileStream(rutaArchivo, FileMode.Open, FileAccess.Read))
            {
                try
                {
                    IWorkbook libro = new XSSFWorkbook(archivo);
                    ISheet hoja = libro.GetSheetAt(0);

                    for (int i = startedRowToRevision; i <= hoja.LastRowNum; i++)
                    {
                        IRow fila = hoja.GetRow(i);
                        if (fila != null)
                        {
                            string referenciaCelda = functions.ObtenerValorCeldaString(fila.GetCell(2)).Trim().ToLower();
                            decimal egresosToCompare = functions.ObtenerValorCeldaDecimal(fila.GetCell(5));
                            ICell fechaCell = fila.GetCell(0);
                            ICell fechaValidacionCell = fila.GetCell(1);
                            string descripcion = functions.ObtenerValorCeldaString(fila.GetCell(3)).Trim();
                            decimal ingresos = functions.ObtenerValorCeldaDecimal(fila.GetCell(4));
                            string numeroFactura = functions.ObtenerValorCeldaString(fila.GetCell(7));
                            string codigoCliente = functions.ObtenerValorCeldaString(fila.GetCell(8));

                            bool coincidenciaUltimosDigitos = false;
                            if (referenciaCelda.EndsWith(digitosBusqueda))
                            {
                                coincidenciaUltimosDigitos = true;
                            }

                            if (coincidenciaUltimosDigitos && ingresos == montoBusqueda && egresosToCompare == 0)
                            {

                                //DateTime fechaValidacion = functions.ObtenerValorCeldaFecha(fechaValidacionCell);
                                //string fechaValidacionFormateada = FormatValidationDate(fechaValidacion);

                                string fechaValidacionFormateada = "";

                                if (CheckCellType(fechaValidacionCell).Equals("Fecha"))
                                {
                                    DateTime fechaValidacion = functions.ObtenerValorCeldaFecha(fechaValidacionCell);
                                    fechaValidacionFormateada = fechaValidacion.ToString("dd/MM/yyyy");
                                }
                                else
                                {
                                    fechaValidacionFormateada = functions.ObtenerValorCeldaString(fechaValidacionCell);
                                }

                                string fechaFormateada = "";

                                if (CheckCellType(fechaCell).Equals("Fecha"))
                                {
                                    DateTime fecha = functions.ObtenerValorCeldaFecha(fechaCell);
                                    fechaFormateada = fecha.ToString("dd/MM/yyyy");
                                }
                                else
                                {
                                    fechaFormateada = functions.ObtenerValorCeldaString(fechaCell);
                                }


                                resultados.Add($"{fechaFormateada}");
                                resultados.Add($"{fechaValidacionFormateada}");
                                resultados.Add($"{referenciaCelda}");
                                resultados.Add($"{descripcion}");
                                resultados.Add($"{ingresos}");
                                resultados.Add($"{egresosToCompare}");
                                resultados.Add($"{numeroFactura}");
                                resultados.Add($"{codigoCliente}");
                                resultados.Add($"{i}");
                            }
                        }
                    }

                    if (resultados.Count == 0)
                    {
                        resultados.Add("No se encontraron coincidencias con la referencia indicada.");
                    }
                }
                catch (Exception ex)
                {
                    resultados.Add("Error al buscar por referencia: " + ex.Message);
                }


                return resultados;
            }
        }

        // Mucha redundancia, hay que mejorarla
        // Por retirar, una vez se valide la calidad de las mejoras implementadas
        public static List<string> SearchByReferenceII(string rutaArchivo, int startedRowToRevision, string referenciaBusqueda)
        {
            List<string> resultados = new List<string>();

            string referenciaBusquedaLower = referenciaBusqueda.Trim().ToLower();

            string DigitosBusqueda = "";

            int digitos = 0;

            if (referenciaBusquedaLower.Length == 6)
            {
                DigitosBusqueda = referenciaBusquedaLower.Substring(referenciaBusquedaLower.Length - 6);
                digitos = 6;
            }
            else if (referenciaBusquedaLower.Length == 4)
            {
                DigitosBusqueda = referenciaBusquedaLower.Substring(referenciaBusquedaLower.Length - 4);
                digitos = 4;
            }
            else
            {
                DigitosBusqueda = referenciaBusquedaLower;
                digitos = referenciaBusquedaLower.Length;
            }


            resultados = forLoopSearchByReferenceII(startedRowToRevision, referenciaBusquedaLower, digitos, rutaArchivo, DigitosBusqueda);

            return resultados;
        }

        // Por retirar, una vez se valide la calidad de las mejoras implementadas
        public static List<string> forLoopSearchByReferenceII(int startedRowToRevision,string referenciaBusquedaLower, int digitos, string rutaArchivo, string digitosBusqueda)
        {
            List<string> resultados = new List<string>();
            ExcelModifyFunctions functions = new ExcelModifyFunctions();

            using (FileStream archivo = new FileStream(rutaArchivo, FileMode.Open, FileAccess.Read))
            {
                if (digitos >= 4 && digitos <= 6)
                {
                    try
                    {
                        IWorkbook libro = new XSSFWorkbook(archivo);
                        ISheet hoja = libro.GetSheetAt(0);

                        for (int i = startedRowToRevision; i <= hoja.LastRowNum; i++)
                        {
                            IRow fila = hoja.GetRow(i);
                            if (fila != null)
                            {
                                string referenciaCelda = functions.ObtenerValorCeldaString(fila.GetCell(2)).Trim().ToLower();
                                decimal egresosToCompare = functions.ObtenerValorCeldaDecimal(fila.GetCell(5));
                                ICell fechaCell = fila.GetCell(0);
                                ICell fechaValidacionCell = fila.GetCell(1);
                                string descripcion = functions.ObtenerValorCeldaString(fila.GetCell(3)).Trim();
                                decimal ingresos = functions.ObtenerValorCeldaDecimal(fila.GetCell(4));
                                string numeroFactura = functions.ObtenerValorCeldaString(fila.GetCell(7));
                                string codigoCliente = functions.ObtenerValorCeldaString(fila.GetCell(8));

                                bool coincidenciaUltimosDigitos = false;
                                if (referenciaCelda.Length >= digitos && referenciaCelda.EndsWith(digitosBusqueda))
                                {
                                    coincidenciaUltimosDigitos = true;
                                }

                                if (coincidenciaUltimosDigitos && egresosToCompare == 0)
                                {

                                    //DateTime fechaValidacion = functions.ObtenerValorCeldaFecha(fechaValidacionCell);
                                    //string fechaValidacionFormateada = FormatValidationDate(fechaValidacion);

                                    string fechaValidacionFormateada = "";
                                   
                                    if (CheckCellType(fechaValidacionCell).Equals("Fecha"))
                                    {
                                        DateTime fechaValidacion = functions.ObtenerValorCeldaFecha(fechaValidacionCell);
                                        fechaValidacionFormateada = fechaValidacion.ToString("dd/MM/yyyy");
                                    }
                                    else
                                    {
                                        fechaValidacionFormateada = functions.ObtenerValorCeldaString(fechaValidacionCell);
                                    }
                                                                                                            
                                    string fechaFormateada = "";

                                    if (CheckCellType(fechaCell).Equals("Fecha"))
                                    {
                                        DateTime fecha = functions.ObtenerValorCeldaFecha(fechaCell);
                                        fechaFormateada = fecha.ToString("dd/MM/yyyy");
                                    }
                                    else
                                    {
                                        fechaFormateada = functions.ObtenerValorCeldaString(fechaCell);
                                    }


                                    resultados.Add($"{fechaFormateada}");
                                    resultados.Add($"{fechaValidacionFormateada}");
                                    resultados.Add($"{referenciaCelda}");
                                    resultados.Add($"{descripcion}");
                                    resultados.Add($"{ingresos}");
                                    resultados.Add($"{egresosToCompare}");
                                    resultados.Add($"{numeroFactura}");
                                    resultados.Add($"{codigoCliente}");
                                    resultados.Add($"{i}");
                                }
                            }
                        }

                        if (resultados.Count == 0)
                        {
                            resultados.Add("No se encontraron coincidencias con la referencia indicada.");
                        }
                    }
                    catch (Exception ex)
                    {
                        resultados.Add("Error al buscar por referencia: " + ex.Message);
                    }
                }
                else
                {
                    resultados = SearchByReference(rutaArchivo, startedRowToRevision, referenciaBusquedaLower);
                }

                return resultados;
            }
        }

       

        public static void ReplaceDataGridViewValues(DataGridView dataGridView1, List<string> myList)
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

        public static void ColorRowsBySecondColumnValue(DataGridView dataGridView)
        {
            if (dataGridView == null || dataGridView.Rows.Count == 0)
            {
                return; // Do nothing if the DataGridView is null or has no rows
            }

            foreach (DataGridViewRow row in dataGridView.Rows)
            {
                // Check if the second column (index 1) has a value
                if (row.Cells.Count > 1 && row.Cells[1].Value != null && !string.IsNullOrEmpty(row.Cells[1].Value.ToString().Trim()))
                {
                    // Change the background color of the entire row
                    foreach (DataGridViewCell cell in row.Cells)
                    {
                        cell.Style.BackColor = Color.Aquamarine;
                    }
                }
                else
                {

                }
            }
        }

        public static void UpdateCellsByRow(string rutaArchivo, int numeroFila, DateTime fecha, DateTime fechaValidacion, string billNumer, string codigoCliente)
        {
            try
            {
                using (FileStream archivo = new FileStream(rutaArchivo, FileMode.Open, FileAccess.ReadWrite))
                {
                    IWorkbook libro = new XSSFWorkbook(archivo);
                    ISheet hoja = libro.GetSheetAt(0);

                    IDataFormat formato = libro.CreateDataFormat();

                    //Estilo con formato general y coloreado Azul

                    IRow fila = hoja.GetRow(numeroFila);
                    if (fila == null)
                        fila = hoja.CreateRow(numeroFila);

                    //Fecha (columna 0)
                    ICell dateCell = fila.GetCell(0) ?? fila.CreateCell(1);
                    dateCell.SetCellValue(fecha.Date);
                    dateCell.CellStyle = CloneSyleAndFormatAndAddColor(libro, dateCell.CellStyle, formato, true);
                    dateCell.CellStyle.DataFormat = formato.GetFormat("dd/MM/yyyy");
                    dateCell.CellStyle.WrapText = true;

                    // Fecha de Validación (columna 1)
                    ICell validationDateCell = fila.GetCell(1) ?? fila.CreateCell(1);
                    validationDateCell.SetCellValue(fechaValidacion.Date);
                    validationDateCell.CellStyle = CloneSyleAndFormatAndAddColor(libro, validationDateCell.CellStyle, formato, true);
                    validationDateCell.CellStyle.DataFormat = formato.GetFormat("dd/MM/yyyy");
                    validationDateCell.CellStyle.WrapText = true;

                    // Número de Factura (columna 7)
                    ICell billCell = fila.GetCell(7) ?? fila.CreateCell(7);
                    billCell.SetCellValue(billNumer);
                    billCell.CellStyle = CloneSyleAndFormatAndAddColor(libro, billCell.CellStyle, formato, true);
                    billCell.CellStyle.WrapText = true;

                    // Código de Cliente (columna 8)
                    ICell clientCell = fila.GetCell(8) ?? fila.CreateCell(8);
                    clientCell.SetCellValue(codigoCliente);
                    clientCell.CellStyle = CloneSyleAndFormatAndAddColor(libro, clientCell.CellStyle, formato, true);
                    clientCell.CellStyle.WrapText = true;

                    //Celdas restantes(columna 6, 5, 4, 3, 2, 0)
                    ICell originalDateCell = fila.GetCell(0) ?? fila.CreateCell(0);
                    originalDateCell.CellStyle = validationDateCell.CellStyle;
                    
                    ICell referenceCell = fila.GetCell(2) ?? fila.CreateCell(2);
                    referenceCell.CellStyle = CloneSyleAndFormatAndAddColor(libro, referenceCell.CellStyle, formato, false);

                    ICell descriptionCell = fila.GetCell(3) ?? fila.CreateCell(3);
                    descriptionCell.CellStyle = CloneSyleAndFormatAndAddColor(libro, descriptionCell.CellStyle, formato, false);

                    ICell incomesCell = fila.GetCell(4) ?? fila.CreateCell(4);
                    incomesCell.CellStyle = CloneSyleAndFormatAndAddColor(libro, incomesCell.CellStyle, formato, false);

                    ICell expensesCell = fila.GetCell(5) ?? fila.CreateCell(5);

                    expensesCell.CellStyle = CloneSyleAndFormatAndAddColor(libro, expensesCell.CellStyle, formato, false);

                    ICell balanceCell = fila.GetCell(6) ?? fila.CreateCell(6);

                    balanceCell.CellStyle = CloneSyleAndFormatAndAddColor(libro, balanceCell.CellStyle, formato, false);


                    //Ajustando el tamaño de la fila para que entren los registros de código de factura
                    // Calcular la altura manualmente *solo si hace falta
                    if (billNumer.Length >= 28 && billNumer.Length <= 58)
                    {
                        int necesaryHeight = 40;
                        fila.HeightInPoints = necesaryHeight;
                    }
                    else if(billNumer.Length >= 59 && billNumer.Length <= 89)
                    {
                        int alturaNecesaria = 60;
                        fila.HeightInPoints = alturaNecesaria;
                    }
                    else if (billNumer.Length >= 90 && billNumer.Length <= 120)
                    {
                        int alturaNecesaria = 80;
                        fila.HeightInPoints = alturaNecesaria;
                    }
                    else if (billNumer.Length >= 121 && billNumer.Length <= 151)
                    {
                        int alturaNecesaria = 100;
                        fila.HeightInPoints = alturaNecesaria;
                    }
                    else if (billNumer.Length >= 152 && billNumer.Length <= 182)
                    {
                        int alturaNecesaria = 120;
                        fila.HeightInPoints = alturaNecesaria;
                    }
                    else if (billNumer.Length >= 183 && billNumer.Length <= 213)
                    {
                        int alturaNecesaria = 140;
                        fila.HeightInPoints = alturaNecesaria;
                    }
                    else if (billNumer.Length >= 214 && billNumer.Length <= 234)
                    {
                        int alturaNecesaria = 160;
                        fila.HeightInPoints = alturaNecesaria;
                    }
                    else if (billNumer.Length >= 235 && billNumer.Length <= 255)
                    {
                        int alturaNecesaria = 180;
                        fila.HeightInPoints = alturaNecesaria;
                    }

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
        private static ICellStyle CloneSyleAndFormatAndAddColor(IWorkbook libro, ICellStyle OriginalStyle, IDataFormat formato, bool generalOrOriginal)
        {
            ICellStyle newStyle = libro.CreateCellStyle();

            if (OriginalStyle != null)
            {
                // Copia los colores de borde
                newStyle.BottomBorderColor = OriginalStyle.BottomBorderColor;
                newStyle.TopBorderColor = OriginalStyle.TopBorderColor;
                newStyle.LeftBorderColor = OriginalStyle.LeftBorderColor;
                newStyle.RightBorderColor = OriginalStyle.RightBorderColor;

                // Copia alineación, fuente, etc.
                newStyle.Alignment = OriginalStyle.Alignment;
                newStyle.VerticalAlignment = OriginalStyle.VerticalAlignment;
                newStyle.WrapText = OriginalStyle.WrapText;
                newStyle.FillBackgroundColor = OriginalStyle.FillBackgroundColor;
                newStyle.ShrinkToFit = OriginalStyle.ShrinkToFit;
                newStyle.Indention = OriginalStyle.Indention;
                newStyle.Rotation = OriginalStyle.Rotation;

                // Sobreescribiendo sombreado al color necesario
                newStyle.FillPattern = FillPattern.SolidForeground;
                newStyle.FillForegroundColor = IndexedColors.LightBlue.Index;

                // Copia los bordes
                newStyle.BorderBottom = OriginalStyle.BorderBottom;
                newStyle.BorderTop = OriginalStyle.BorderTop;
                newStyle.BorderLeft = OriginalStyle.BorderLeft;
                newStyle.BorderRight = OriginalStyle.BorderRight;
            }

            // Asigna el formato "General" o el formato original
            newStyle.DataFormat = generalOrOriginal ? formato.GetFormat("General") : OriginalStyle.DataFormat;

            return newStyle;
        }

        public static string CheckCellType(ICell celda)
        {
            if (celda == null)
            {
                return "Celda vacía";
            }

            switch (celda.CellType)
            {
                case CellType.String:
                    return "Texto";

                case CellType.Numeric:
                    if (DateUtil.IsCellDateFormatted(celda))
                    {
                        return "Fecha";
                    }
                    else
                    {
                        return "Número";
                    }

                case CellType.Boolean:
                    return "Booleano";

                case CellType.Formula:
                    return "Fórmula";

                case CellType.Blank:
                    return "Celda vacía";

                default:
                    return "Tipo de dato desconocido";
            }
        }

        public static string CopyExcelFile(string rutaArchivoOrigen, string rutaDirectorioDestino)
        {
            try
            {
                // Obtener la información del archivo original
                FileInfo archivoOrigen = new FileInfo(rutaArchivoOrigen);
                string nombreArchivo = Path.GetFileNameWithoutExtension(rutaArchivoOrigen);
                string extensionArchivo = archivoOrigen.Extension;

                // Obtener la fecha y hora actual y formatearla para el nombre del archivo
                string fechaHoraActual = DateTime.Now.ToString("yyyyMMdd_HHmmss");

                // Eliminar caracteres no válidos para nombres de archivo en Windows
                string nombreArchivoSeguro = Regex.Replace(nombreArchivo, @"[<>:""/\\|?*]", "_");

                // Construir el nombre del archivo de destino
                string nombreArchivoDestino = $"{nombreArchivoSeguro}_{fechaHoraActual}{extensionArchivo}";
                string rutaArchivoDestino = Path.Combine(rutaDirectorioDestino, nombreArchivoDestino);

                // Copiar el archivo
                File.Copy(rutaArchivoOrigen, rutaArchivoDestino);

                Console.WriteLine($"Archivo copiado exitosamente a: {rutaArchivoDestino}");
                return rutaArchivoDestino;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error al crear la copia del archivo: {ex.Message}");
                return null;
            }
        }

        public static bool isRowFilledwithColor(string rutaArchivo, int numeroFila, int cantidadCeldasARevisar)
        {
            try
            {
                using (FileStream archivo = new FileStream(rutaArchivo, FileMode.Open, FileAccess.Read))
                {
                    XSSFWorkbook libro = new XSSFWorkbook(archivo);
                    ISheet hoja = libro.GetSheetAt(0);
                    IRow fila = hoja.GetRow(numeroFila); // Las filas en NPOI son 0-based

                    if (fila != null)
                    {
                        for (int i = 0; i < cantidadCeldasARevisar; i++)
                        {
                            ICell celda = fila.GetCell(i);
                            if (celda != null)
                            {
                                ICellStyle estilo = celda.CellStyle;
                                if (estilo != null)
                                {
                                    FillPattern fillPattern = estilo.FillPattern;

                                    // Si hay un patrón de relleno distinto a "ninguno", consideramos que tiene color
                                    if (fillPattern != FillPattern.NoFill)
                                    {
                                        return true;
                                    }
                                    else
                                    {
                                        // Si no hay patrón, verificamos el color de fondo
                                        XSSFColor backgroundColor = (XSSFColor)estilo.FillForegroundColorColor;
                                        if (backgroundColor != null)
                                        {
                                            // Comparamos con el color de fondo predeterminado (puede variar, pero blanco es común)
                                            byte[] defaultBackgroundRGB = new byte[] { 255, 255, 255 };
                                            byte[] backgroundColorRGB = backgroundColor.RGB;

                                            // Si el color RGB no es nulo y no es igual al blanco predeterminado
                                            if (backgroundColorRGB != null && !backgroundColorRGB.SequenceEqual(defaultBackgroundRGB))
                                            {
                                                return true;
                                            }
                                            else if (backgroundColor.Indexed != HSSFColor.Automatic.Index)
                                            {
                                                // Verificamos si el índice del color no es el automático
                                                return true;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }

                    return false;
                }
            }
            catch (FileNotFoundException)
            {
                return false;
            }
            catch (IOException ex)
            {
                return false;
            }
            return false;
        }

    }
}

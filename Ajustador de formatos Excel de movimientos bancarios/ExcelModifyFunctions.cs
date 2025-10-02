using NPOI.HSSF.UserModel;
using NPOI.OpenXmlFormats;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using SixLabors.ImageSharp.Formats.Gif;
using System;
using System.Collections.Generic;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.Linq;
using System.Security.Cryptography.Xml;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;



namespace Ajustador_de_formatos_Excel_de_movimientos_bancarios
{
    internal class ExcelModifyFunctions
    {

        FileAccessChecker FileAccessC = new FileAccessChecker();
        BankFormatsFixerFunctions.Exterior bancoExterior = new BankFormatsFixerFunctions.Exterior();
        BankFormatsFixerFunctions.Mercantil bancoMercantil = new BankFormatsFixerFunctions.Mercantil();
        BankFormatsFixerFunctions.Banesco bancoBanesco = new BankFormatsFixerFunctions .Banesco();
        BankFormatsFixerFunctions.BDV bancoVenezuela = new BankFormatsFixerFunctions.BDV ();

        public void AttachExcelFile(ComboBox BankSelector, TextBox ExcelFilePath)
        {

            if (FileAccessC.IsOpen(ExcelFilePath.Text))
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


                        if (bancoBanesco.bankValidator(ExcelFilePath.Text) == 1)
                        {

                            bancoBanesco.fixFormat(ExcelFilePath);

                        }
                        else if (bancoBanesco.bankValidator(ExcelFilePath.Text) == 2)
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
                        if (bancoVenezuela.bankValidator(ExcelFilePath.Text) == 1)
                        {
                            bancoVenezuela.fixFormat(ExcelFilePath);
                        }
                        else if (bancoVenezuela.bankValidator(ExcelFilePath.Text) == 2)
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
                        if (bancoMercantil.bankValidator(ExcelFilePath.Text) == 1)
                        {

                            bancoMercantil.fixFormat(ExcelFilePath);

                        }
                        else if (bancoMercantil.bankValidator(ExcelFilePath.Text) == 2)
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
                        if (bancoExterior.BankValidator(ExcelFilePath.Text) == 1)
                        {

                            bancoExterior.fixFormat(ExcelFilePath);

                        }
                        else if (bancoExterior.BankValidator(ExcelFilePath.Text) == 2)
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
                                    string valorString = getValueCellString(celdaOrigen);

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
                                    string valorString = getValueCellString(celdaOrigen);

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

        // POR ELIMINAR -----------------------------------------------------------------------
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

        // ----------------------------------------------------------------------------------------------

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
                            string referencia = getValueCellString(fila.GetCell(2)).Trim().ToLower(); // Normalizar referencia
                            string descripcion = getValueCellString(fila.GetCell(3)).Trim().ToLower(); // Normalizar descripción
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
                            string fecha = getValueCellString(fila.GetCell(0));
                            string referencia = getValueCellString(fila.GetCell(1));
                            string descripcion = getValueCellString(fila.GetCell(2)).Trim().ToLower(); // Normalizar referencia
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

        public static string getValueCellString(ICell celda)
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
        
        // Función auxiliar para copiar el contenido y formato de una celda

        public ICellStyle CopyCellStyle(ICellStyle estiloOrigen, IWorkbook libro)
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

        // POR ELIMINAR --------------------------------------------------------------
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

        // ----------------------------------------------------------------------------------------

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

       

        // POR ELIMINAR ----------------------------------------------------------------
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


        // ----------------------------------------------------------------------------------------------
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

        public List<string> CopyDateColumnsAsStrings(string rutaArchivo, int sheetName, int columna)
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
                                string valorCelda = getCellValueAsStringII(celda);
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

        public string getCellValueAsStringII(ICell celda)
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

        public void changeCellTextFromListInReverseOrder(string rutaArchivo, int columna, int nHoja, List<string> listaColumna1)
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

        public void changeCellTextFromListInTheSameOrder(string rutaArchivo, int columna, int nHoja, List<string> listaColumna1)
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

        public void replaceEmptyCellsWithZero(string rutaArchivo, int sheetName, int columnaIndex)
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
                            else if (getValueCellString(celda).Equals("") || getCellValueAsStringII(celda) == "" || getCellValueAsStringII(celda).Equals(""))
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


    }

}
    





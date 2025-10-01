using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace Ajustador_de_formatos_Excel_de_movimientos_bancarios
{
    public partial class ExcelFixForm2 : Form
    {

        string rutaArchivoSeleccionado = "";
        FileAccessChecker FileAccessC = new FileAccessChecker();

        public ExcelFixForm2()
        {
            InitializeComponent();
        }

        private void SearchBankMovesProcess()
        {


            ExcelModifyFunctions functions = new ExcelModifyFunctions();

            if (!string.IsNullOrEmpty(rutaArchivoSeleccionado))
            {
                //Este if verifica si el archivo está abierto o no. (No se puede procesar estando abierto)
                if (FileAccessC.IsOpen(rutaArchivoSeleccionado))
                {
                    List<string> myList = new List<string>();

                    if (string.IsNullOrWhiteSpace(ReferenceNumberTextBox.Text) && string.IsNullOrWhiteSpace(bankMovMountTextBox.Text))
                    {

                        MessageBox.Show("Ambos campos no pueden estar vacíos.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;

                    }

                    else if (string.IsNullOrWhiteSpace(ReferenceNumberTextBox.Text) && !string.IsNullOrWhiteSpace(bankMovMountTextBox.Text))
                    {

                        myList = MovValidatorFunctions.SearchByMount(rutaArchivoSeleccionado, 0, decimal.Parse(bankMovMountTextBox.Text));

                    }

                    else if (!string.IsNullOrWhiteSpace(ReferenceNumberTextBox.Text) && string.IsNullOrWhiteSpace(bankMovMountTextBox.Text))
                    {

                        myList = MovValidatorFunctions.SearchByReferenceIII(rutaArchivoSeleccionado, 0, ReferenceNumberTextBox.Text);

                    }

                    else
                    {

                        // Filtrar por monto y referencia, actualizado para considerar referencias cortas                   

                        myList = MovValidatorFunctions.SearchByReferenceandAmountII(rutaArchivoSeleccionado, 0, ReferenceNumberTextBox.Text, decimal.Parse(bankMovMountTextBox.Text));
                   

                    }

                    
                    MovValidatorFunctions.ReplaceDataGridViewValues(dataGridView1, myList);
                    MovValidatorFunctions.ColorRowsBySecondColumnValue(dataGridView1);
                }


            }

            else
            {
                MessageBox.Show("Primero debes adjuntar un archivo Excel.", "Archivo no seleccionado", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }


        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Crear una instancia del segundo formulario
            PrincipalMenu Option1Form = new PrincipalMenu();

            // Mostrar el segundo formulario
            Option1Form.Show();

            // Cerrar el formulario actual (Form1)
            this.Hide();

        }

        private void SearchMoveButton_Click(object sender, EventArgs e)
        {

            SearchBankMovesProcess();

        }

        private void dateTitleLabel_Click(object sender, EventArgs e)
        {

        }

        private void ExcelFixForm2_Load(object sender, EventArgs e)
        {

        }

        private void FormMovValidator_FormClosing(object sender, FormClosingEventArgs e)
        {
            Application.Exit(); // Cierra toda la aplicación
        }

        private void ReferenceNumberTextBox_TextChanged(object sender, EventArgs e)
        {

        }

        private void attachFilebutton_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();

            // Filtrar los tipos de archivos que se pueden seleccionar (opcional)
            openFileDialog.Filter = "Archivos de Excel (*.xls;*.xlsx;*.xlsm;*.xlsb;*.xltx)|*.xls;*.xlsx;*.xlsm;*.xlsb;*.xltx|Todos los archivos (*.*)|*.*";
            openFileDialog.Title = "Selecciona un archivo de Excel";

            // Mostrar el cuadro de diálogo y obtener el resultado
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                // El usuario ha seleccionado un archivo
                rutaArchivoSeleccionado = openFileDialog.FileName;
                string fileName = Path.GetFileNameWithoutExtension(rutaArchivoSeleccionado);

                // Mostrar la ruta del archivo en un cuadro de texto o donde sea necesario
                FileNameTextBox.Text = fileName;

            }
            else
            {
                // El usuario ha cancelado la selección del archivo
                MessageBox.Show("No se ha seleccionado ningún archivo.");
            }
        }

        private void ReferenceNumberTextBox_TextChanged_1(object sender, EventArgs e)
        {

        }

        private void ReferenceNumberTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Verifica si la tecla presionada es un dígito o una tecla de control (como borrar)
            if (!char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar))
            {
                // Si no es un dígito ni una tecla de control, ignora la tecla presionada
                e.Handled = true;
            }
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void updateRowButton_Click(object sender, EventArgs e)
        {

            ExcelModifyFunctions functions = new ExcelModifyFunctions();

            if (!string.IsNullOrEmpty(rutaArchivoSeleccionado))
            {
                //Este if verifica si el archivo está abierto o no. (No se puede procesar estando abierto)
                if (FileAccessC.IsOpen(rutaArchivoSeleccionado))
                {

                    if (dataGridView1.SelectedRows.Count > 0)
                    {

                        // Continuar con el flujo porque hay una fila seleccionada
                        DataGridViewRow filaSeleccionada = dataGridView1.SelectedRows[0];

                        string valor = filaSeleccionada.Cells[8].Value?.ToString();
                        string validationDateDataGridView = filaSeleccionada.Cells[1].Value?.ToString();

                        
                        if (!string.IsNullOrWhiteSpace(billCodeTextBox.Text) && !string.IsNullOrWhiteSpace(clientCodeTextBox.Text))
                        {
                            if (MovValidatorFunctions.isRowFilledwithColor(rutaArchivoSeleccionado, int.Parse(valor), 4) || !string.IsNullOrEmpty(validationDateDataGridView))
                            {
                                
                                MessageBox.Show("Este movimiento ya ha sido validado el " + validationDateDataGridView, "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                            }
                            else
                            {
                                DateTime fechaSeleccionada;

                                try
                                {
                                    fechaSeleccionada = Convert.ToDateTime(filaSeleccionada.Cells[0].Value);
                                    // Ahora 'fechaSeleccionada' es de tipo DateTime y contiene la fecha del DataGridView.
                                }
                                catch (FormatException ex)
                                {
                                    // Manejar el caso en que el valor no es una fecha válida.
                                    MessageBox.Show($"El valor '{filaSeleccionada.Cells[0].Value}' no tiene un formato de fecha válido.", "Error de formato");
                                    fechaSeleccionada = DateTime.MinValue; // O asigna otro valor por defecto.
                                }

                                MovValidatorFunctions.UpdateCellsByRow(rutaArchivoSeleccionado, int.Parse(valor), fechaSeleccionada, DateTime.Now, billCodeTextBox.Text, clientCodeTextBox.Text);
                                
                                SearchBankMovesProcess();
                                                                
                            }

                        }
                        else if (string.IsNullOrWhiteSpace(clientCodeTextBox.Text))
                        {
                            MessageBox.Show("Debe llenar el campo: Código de cliente", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                        else if (string.IsNullOrWhiteSpace(billCodeTextBox.Text))
                        {
                            MessageBox.Show("Debe llenar el campo: Código de facturas", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }


                    }
                    else
                    {
                        // Mostrar mensaje de advertencia
                        MessageBox.Show("Debe seleccionar una fila antes de continuar.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }


                }

            }

            else
            {
                MessageBox.Show("Primero debes adjuntar un archivo Excel.", "Archivo no seleccionado", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

        }

        private void clientCodeTextBox_TextChanged(object sender, EventArgs e)
        {

        }

        private void bankMovMountTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Permitir dígitos, el punto decimal y teclas de control
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != ','))
            {
                e.Handled = true; // Ignorar el carácter ingresado
            }

            // Permitir solo un punto decimal
            if ((e.KeyChar == ',') && (bankMovMountTextBox.Text.IndexOf(',') > -1 || string.IsNullOrEmpty(bankMovMountTextBox.Text)))
            {
                e.Handled = true; // Ignorar el carácter ingresado
            }

            // Limitar la longitud del texto a 50 caracteres
            if (bankMovMountTextBox.Text.Length >= 50 && !char.IsControl(e.KeyChar))
            {
                e.Handled = true; // Ignorar el carácter ingresado
            }

        }

        private void bankMovMountTextBox_TextChanged(object sender, EventArgs e)
        {

        }

        private void billCodeTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Limitar la longitud del texto a 250 caracteres
            if (bankMovMountTextBox.Text.Length >= 250 && !char.IsControl(e.KeyChar))
            {
                e.Handled = true; // Ignorar el carácter ingresado
            }
        }

        private void clientCodeTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Limitar la longitud del texto a 250 caracteres
            if (bankMovMountTextBox.Text.Length >= 250 && !char.IsControl(e.KeyChar))
            {
                e.Handled = true; // Ignorar el carácter ingresado
            }
        }

        private void billCodeTextBox_TextChanged(object sender, EventArgs e)
        {

        }

        private void createCopy_Click(object sender, EventArgs e)
        {           
            string directorioDestino = getDestinationPathForCopy();

            if (directorioDestino != null)
            {
                // Usar el directorio de destino para crear la copia del archivo Excel

                if (!string.IsNullOrEmpty(rutaArchivoSeleccionado))
                {

                    string rutaArchivoOrigen = rutaArchivoSeleccionado; // Reemplaza con la ruta de tu archivo Excel original

                    string rutaArchivoDestino = getDestinationPathForCopy();

                    if (rutaArchivoDestino != null)
                    {
                        MovValidatorFunctions.CopyExcelFile(rutaArchivoOrigen, rutaArchivoDestino);
                        MessageBox.Show($"Archivo copiado exitosamente a: {rutaArchivoDestino}", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show("Error al crear la copia del archivo.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                }

                else
                {

                    MessageBox.Show("Primero debes adjuntar un archivo Excel.", "Archivo no seleccionado", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                }
                
            }

            else
            {

                MessageBox.Show("Debes elegir una dirección válida para la copia.", "Archivo no seleccionado", MessageBoxButtons.OK, MessageBoxIcon.Warning);

            }
        }

        public static string getDestinationPathForCopy()
        {
            using (FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog())
            {
                DialogResult resultado = folderBrowserDialog.ShowDialog();

                if (resultado == DialogResult.OK && !string.IsNullOrWhiteSpace(folderBrowserDialog.SelectedPath))
                {
                    return folderBrowserDialog.SelectedPath;
                }
                else
                {
                    MessageBox.Show("Destino no válido o no seleccionado.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return null;
                }
            }
        }
    }
}

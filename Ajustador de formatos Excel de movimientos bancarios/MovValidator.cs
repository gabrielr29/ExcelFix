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

namespace Ajustador_de_formatos_Excel_de_movimientos_bancarios
{
    public partial class ExcelFixForm2 : Form
    {

        string rutaArchivoSeleccionado = "";

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
                if (!functions.IsOpen(rutaArchivoSeleccionado))
                {
                    List<string> myList = new List<string>();

                    if (string.IsNullOrWhiteSpace(ReferenceNumberTextBox.Text) && string.IsNullOrWhiteSpace(bankMovMountTextBox.Text))
                    {

                        MessageBox.Show("Ambos campos no pueden estar vacío.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;

                    }

                    else if (string.IsNullOrWhiteSpace(ReferenceNumberTextBox.Text) && !string.IsNullOrWhiteSpace(bankMovMountTextBox.Text))
                    {

                        myList = MovValidatorFunctions.SearchByMount(rutaArchivoSeleccionado, 0, decimal.Parse(bankMovMountTextBox.Text));

                    }

                    else if (!string.IsNullOrWhiteSpace(ReferenceNumberTextBox.Text) && string.IsNullOrWhiteSpace(bankMovMountTextBox.Text))
                    {

                        myList = MovValidatorFunctions.SearchByReference(rutaArchivoSeleccionado, 0, ReferenceNumberTextBox.Text);

                    }

                    else
                    {

                        // Lógica para cargar y filtrar el Excel usando la ruta                    

                        myList = MovValidatorFunctions.SearchByReferenceAndMount(rutaArchivoSeleccionado, 0, ReferenceNumberTextBox.Text, decimal.Parse(bankMovMountTextBox.Text));


                    }


                    MovValidatorFunctions.ReplaceDataGridViewValues2(dataGridView1, myList);

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

        private void button2_Click(object sender, EventArgs e)
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

                // Opcional: Puedes hacer algo con la ruta del archivo aquí
                // Por ejemplo, puedes leer el contenido del archivo o procesarlo de alguna manera
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
                if (!functions.IsOpen(rutaArchivoSeleccionado))
                {

                    if (dataGridView1.SelectedRows.Count > 0)
                    {

                        // Continuar con el flujo porque hay una fila seleccionada
                        DataGridViewRow filaSeleccionada = dataGridView1.SelectedRows[0];

                        string valor = filaSeleccionada.Cells[8].Value?.ToString();


                        if (!string.IsNullOrWhiteSpace(billCodeTextBox.Text) && !string.IsNullOrWhiteSpace(billCodeTextBox.Text))
                        {
                            MovValidatorFunctions.UpdateCellsByRow(rutaArchivoSeleccionado, int.Parse(valor), DateTime.Now, billCodeTextBox.Text, clientCodeTextBox.Text);
                            SearchBankMovesProcess();
                        }
                        else if (!string.IsNullOrWhiteSpace(billCodeTextBox.Text))
                        {
                            MessageBox.Show("Debe llenar el campo: Código de cliente", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                        else if (!string.IsNullOrWhiteSpace(billCodeTextBox.Text))
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
    }
}

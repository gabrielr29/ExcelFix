namespace Ajustador_de_formatos_Excel_de_movimientos_bancarios
{
    public partial class Form1 : Form
    {

        ExcelModifyFunctions functions = new ExcelModifyFunctions();
        public Form1()
        {
            InitializeComponent();
        }

        private void Executebutton_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBoxPath.Text))
            {
                functions.AttachExcelFile(bankSelector, textBoxPath);
            }
            else
            {
               MessageBox.Show("Primero debes adjuntar un archivo Excel.", "Archivo no seleccionado", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();

            // Filtrar los tipos de archivos que se pueden seleccionar (opcional)
            openFileDialog.Filter = "Archivos de Excel (*.xls;*.xlsx;*.xlsm;*.xlsb;*.xltx)|*.xls;*.xlsx;*.xlsm;*.xlsb;*.xltx|Todos los archivos (*.*)|*.*";
            openFileDialog.Title = "Selecciona un archivo de Excel";

            // Mostrar el cuadro de diálogo y obtener el resultado
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                // El usuario ha seleccionado un archivo
                string rutaArchivo = openFileDialog.FileName;
                string fileName = Path.GetFileNameWithoutExtension(rutaArchivo);

                // Mostrar la ruta del archivo en un cuadro de texto o donde sea necesario
                textBoxPath.Text = rutaArchivo;
                textBoxFileName.Text = fileName;

            }
            else
            {
                // El usuario ha cancelado la selección del archivo
                MessageBox.Show("No se ha seleccionado ningún archivo.");
            }



        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBoxFileName_TextChanged(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            // Crear una instancia
            PrincipalMenu OptionForm = new PrincipalMenu();

            // Mostrar 
            OptionForm.Show();

            // Cerrar el formulario actual 
            this.Hide();
        }
        private void FormBankFormatFixer_FormClosing(object sender, FormClosingEventArgs e)
        {
            Application.Exit(); // Cierra toda la aplicación
        }
    }



    }

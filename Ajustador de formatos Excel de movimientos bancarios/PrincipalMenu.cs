using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms;


namespace Ajustador_de_formatos_Excel_de_movimientos_bancarios
{
    public partial class PrincipalMenu : Form
    {
        public PrincipalMenu()
        {
            InitializeComponent();
        }

        private void Form3_Load(object sender, EventArgs e)
        {
            int currentYear = DateTime.Now.Year;
            lblCopyright.Text = $"© {currentYear} ExcelFix · Desarrollado por Gabriel Rodriguez";
        }

        private void AjustBankFormats_Click(object sender, EventArgs e)
        {
            // Crear una instancia del segundo formulario
            Form1 Option1Form = new Form1();

            // Mostrar el segundo formulario
            Option1Form.Show();

            // Cerrar el formulario actual (Form1)
            this.Hide();
        }

        private void MovValidator_Click(object sender, EventArgs e)
        {

            // Crear una instancia del segundo formulario
            ExcelFixForm2 Option1Form = new ExcelFixForm2();

            // Mostrar el segundo formulario
            Option1Form.Show();

            // Cerrar el formulario actual (Form1)
            this.Hide();

        }

        private void FAQForm_Click(object sender, EventArgs e)
        {
            // Crear una instancia del segundo formulario
             FAQ Option1Form = new FAQ();

            // Mostrar el segundo formulario
            Option1Form.Show();

            // Cerrar el formulario actual (Form1)
            this.Hide();


        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void FormPrincipalMenu_FormClosing(object sender, FormClosingEventArgs e)
        {
            Application.Exit(); // Cierra toda la aplicación
        }
    }




}

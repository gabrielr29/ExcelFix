using System;
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
    public partial class FAQ : Form
    {
        public FAQ()
        {
            InitializeComponent();
        }

        private void kryptonRichTextBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

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

        private void FAQ_Load(object sender, EventArgs e)
        {

        }

        private void FormFAQ_FormClosing(object sender, FormClosingEventArgs e)
        {
            Application.Exit(); // Cierra toda la aplicación
        }
    }
}

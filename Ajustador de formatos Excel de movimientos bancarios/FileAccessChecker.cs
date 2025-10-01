using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Ajustador_de_formatos_Excel_de_movimientos_bancarios
{
    internal class FileAccessChecker
    {

        // Función auxiliar para verificar si la excepción es por archivo bloqueado
        public virtual bool IsFileLocked(IOException e)
        {
            return e.HResult == -2147024864; // Error HRESULT: 0x80070020
        }

        public bool IsOpen(string path)
        {
            try
            {
                using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.None))
                {
                    // El archivo no está abierto por otro proceso
                    return false;
                }
            }
            catch (IOException ex)
            {
                // El archivo está abierto por otro proceso
                if (IsFileLocked(ex))
                {
                    MessageBox.Show("El archivo está abierto. Cierre el archivo para continuar.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    // Otro error de E/S
                    MessageBox.Show("Error al acceder al archivo: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return true;
            }
        }



    }
}

using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Ajustador_de_formatos_Excel_de_movimientos_bancarios.BusinessLogic
{
    internal class NPOIStyleHelper
    {
        // Caché estático para estilos de celda
        private static readonly Dictionary<string, ICellStyle> _cellStyleCache = new Dictionary<string, ICellStyle>();

        // Caché estático para fuentes. Es crucial gestionarlas también.
        private static readonly Dictionary<string, IFont> _fontCache = new Dictionary<string, IFont>();

        /// <summary>
        /// Clona un estilo existente o crea uno nuevo, reutilizando estilos de un caché.
        /// Solo considera las propiedades explícitamente listadas para la clonación y la clave del caché.
        /// </summary>
        /// <param name="libro">El libro de trabajo (IWorkbook) al que pertenece el estilo.</param>
        /// <param name="originalStyle">El estilo de celda original para clonar propiedades.</param>
        /// <param name="formato">El IDataFormat del libro, usado para obtener formatos de datos.</param>
        /// <param name="generalOrOriginal">Si es true, aplica formato "General"; de lo contrario, aplica el formato del estilo original.</param>
        /// <returns>Un ICellStyle reutilizado o recién creado con las propiedades especificadas.</returns>
        public static ICellStyle CloneSyleAndFormatAndAddColor(IWorkbook libro, ICellStyle originalStyle, IDataFormat formato, bool generalOrOriginal)
        {
            // 1. Construir una clave única para el estilo deseado
            StringBuilder cacheKeyBuilder = new StringBuilder();

            // Incluir el booleano 'generalOrOriginal' y el color fijo 'LightBlue'
            cacheKeyBuilder.Append($"GOO:{generalOrOriginal};FGC:LightBlue;FP:{FillPattern.SolidForeground};");

            // Incluir propiedades del estilo original si existe
            if (originalStyle != null)
            {
                // Propiedades de bordes
                cacheKeyBuilder.Append($"BB{originalStyle.BorderBottom}BC{originalStyle.BottomBorderColor};");
                cacheKeyBuilder.Append($"TB{originalStyle.BorderTop}TC{originalStyle.TopBorderColor};");
                cacheKeyBuilder.Append($"LB{originalStyle.BorderLeft}LC{originalStyle.LeftBorderColor};");
                cacheKeyBuilder.Append($"RB{originalStyle.BorderRight}RC{originalStyle.RightBorderColor};");

                // Alineación y otras propiedades
                cacheKeyBuilder.Append($"AL{originalStyle.Alignment};VA{originalStyle.VerticalAlignment};WT{originalStyle.WrapText};");
                cacheKeyBuilder.Append($"FBC{originalStyle.FillBackgroundColor};SF{originalStyle.ShrinkToFit};IND{originalStyle.Indention};ROT{originalStyle.Rotation};");

                // Formato de datos
                // Obtener el formato como string para que la clave sea robusta
                string dataFormatString = formato.GetFormat(originalStyle.DataFormat);
                cacheKeyBuilder.Append($"DF:{dataFormatString};");

                // Clave de la fuente (gestión de la fuente es crucial)
                IFont originalFont = originalStyle.GetFont(libro);
                cacheKeyBuilder.Append($"FONT_KEY:{GetFontCacheKey(originalFont)};");
            }
            else
            {
                cacheKeyBuilder.Append("NO_ORIGINAL_STYLE;");
                // Para el caso sin originalStyle, el formato por defecto será "General"
                cacheKeyBuilder.Append($"DF:{formato.GetFormat("General")};");
                // Y la fuente por defecto será la del libro
                cacheKeyBuilder.Append($"FONT_KEY:{GetFontCacheKey(libro.GetFontAt(0))};"); // Generalmente la primera fuente es la por defecto
            }

            string cacheKey = cacheKeyBuilder.ToString();

            // 2. Intentar recuperar el estilo del caché
            if (_cellStyleCache.TryGetValue(cacheKey, out ICellStyle existingStyle))
            {
                return existingStyle;
            }

            // 3. Si no existe en el caché, crear el nuevo estilo
            ICellStyle newStyle = libro.CreateCellStyle();

            if (originalStyle != null)
            {
                // Usar CloneStyleFrom para copiar las propiedades del estilo original.
                // Esto es más conciso y copia la mayoría de las propiedades correctamente.
                newStyle.CloneStyleFrom(originalStyle);

                // AHORA sobrescribimos las propiedades que tu función siempre aplica o modifica.
                // (Tu función original sobrescribe el sombreado y establece un color fijo)
                newStyle.FillPattern = FillPattern.SolidForeground;
                newStyle.FillForegroundColor = IndexedColors.LightBlue.Index;

                // Gestionar la fuente clonando y almacenando en caché
                IFont originalFont = originalStyle.GetFont(libro);
                if (originalFont != null)
                {
                    newStyle.SetFont(GetOrCreateFont(libro, originalFont));
                }
            }
            else // Caso donde no hay estilo original, solo aplicar las propiedades base que siempre se establecen.
            {
                // Establecer propiedades por defecto si no hay un estilo original para clonar
                newStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.General; // Usar el using correcto para NPOI
                newStyle.VerticalAlignment = VerticalAlignment.Bottom; // Usar el using correcto para NPOI
                newStyle.WrapText = false;
                newStyle.ShrinkToFit = false;
                newStyle.Indention = 0;
                newStyle.Rotation = 0;

                // Y las propiedades de sombreado que tu función siempre aplica
                newStyle.FillPattern = FillPattern.SolidForeground;
                newStyle.FillForegroundColor = IndexedColors.LightBlue.Index;

                // Establecer la fuente por defecto del libro
                newStyle.SetFont(GetOrCreateFont(libro, libro.GetFontAt(0))); // Asumiendo que la fuente en índice 0 es la por defecto
            }

            // Asigna el formato "General" o el formato del estilo original
            newStyle.DataFormat = generalOrOriginal
                                  ? formato.GetFormat("General")
                                  : originalStyle != null ? originalStyle.DataFormat : formato.GetFormat("General"); // Default a "General" si no hay original

            // 4. Guardar el nuevo estilo en el caché antes de retornarlo
            _cellStyleCache[cacheKey] = newStyle;

            return newStyle;
        }

        /// <summary>
        /// Genera una clave única para una fuente basada en sus propiedades.
        /// </summary>
        private static string GetFontCacheKey(IFont font)
        {
            if (font == null) return "NULL_FONT";

            // Incluye las propiedades de la fuente que la hacen única.
            // Asegúrate de incluir todas las propiedades que son relevantes para tu aplicación.
            StringBuilder fontKeyBuilder = new StringBuilder();
            fontKeyBuilder.Append($"BW{font.IsBold};C{font.Color};HS{font.FontHeightInPoints};FN{font.FontName};I{font.IsItalic};S{font.IsStrikeout};U{font.Underline};");
            // Nota: 'Boldweight' está obsoleto, es mejor usar 'IsBold'
            return fontKeyBuilder.ToString();
        }

        /// <summary>
        /// Reutiliza o crea una fuente, gestionándola en un caché.
        /// </summary>
        private static IFont GetOrCreateFont(IWorkbook libro, IFont originalFont)
        {
            if (originalFont == null)
            {
                // Podrías retornar la fuente por defecto del libro si originalFont es null.
                // La fuente en el índice 0 suele ser la fuente por defecto.
                return GetOrCreateFont(libro, libro.GetFontAt(0));
            }

            string fontCacheKey = GetFontCacheKey(originalFont);

            if (_fontCache.TryGetValue(fontCacheKey, out IFont existingFont))
            {
                return existingFont;
            }

            IFont newFont = libro.CreateFont();
            newFont.IsBold = originalFont.IsBold; // Usar IsBold en lugar de Boldweight
            newFont.Color = originalFont.Color;
            newFont.FontHeightInPoints = originalFont.FontHeightInPoints;
            newFont.FontName = originalFont.FontName;
            newFont.IsItalic = originalFont.IsItalic;
            newFont.IsStrikeout = originalFont.IsStrikeout;
            newFont.Underline = originalFont.Underline;
            // Copia cualquier otra propiedad de fuente que uses.

            _fontCache[fontCacheKey] = newFont;
            return newFont;
        }
    }





}


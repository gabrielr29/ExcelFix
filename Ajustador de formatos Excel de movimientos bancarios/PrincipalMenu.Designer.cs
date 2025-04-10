namespace Ajustador_de_formatos_Excel_de_movimientos_bancarios
{
    partial class PrincipalMenu
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(PrincipalMenu));
            label1 = new Label();
            label2 = new Label();
            FixBankFormatsButton = new Krypton.Toolkit.KryptonButton();
            MovValidatorButton = new Krypton.Toolkit.KryptonButton();
            FAQButton = new Krypton.Toolkit.KryptonButton();
            label3 = new Label();
            SuspendLayout();
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Font = new Font("Consolas", 18F, FontStyle.Italic, GraphicsUnit.Point, 0);
            label1.Location = new Point(98, 33);
            label1.Name = "label1";
            label1.Size = new Size(285, 28);
            label1.TabIndex = 0;
            label1.Text = "Bienvenido a ExcelFix";
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Font = new Font("Consolas", 14.25F, FontStyle.Italic, GraphicsUnit.Point, 0);
            label2.Location = new Point(120, 61);
            label2.Name = "label2";
            label2.Size = new Size(230, 22);
            label2.TabIndex = 5;
            label2.Text = "¿Qué deseas hacer hoy?";
            // 
            // FixBankFormatsButton
            // 
            FixBankFormatsButton.Location = new Point(98, 111);
            FixBankFormatsButton.Name = "FixBankFormatsButton";
            FixBankFormatsButton.Size = new Size(285, 57);
            FixBankFormatsButton.StateCommon.Back.Color1 = Color.FromArgb(240, 244, 250);
            FixBankFormatsButton.StateCommon.Back.Color2 = Color.FromArgb(240, 244, 250);
            FixBankFormatsButton.StateCommon.Border.Color1 = Color.FromArgb(202, 215, 234);
            FixBankFormatsButton.StateCommon.Border.Rounding = 15F;
            FixBankFormatsButton.StateCommon.Border.Width = 1;
            FixBankFormatsButton.StateCommon.Content.ShortText.Color1 = Color.Black;
            FixBankFormatsButton.StateCommon.Content.ShortText.Font = new Font("Consolas", 10F, FontStyle.Bold);
            FixBankFormatsButton.StateTracking.Back.Color1 = Color.FromArgb(220, 230, 245);
            FixBankFormatsButton.StateTracking.Border.Color1 = Color.FromArgb(184, 204, 230);
            FixBankFormatsButton.TabIndex = 6;
            FixBankFormatsButton.TabStop = false;
            FixBankFormatsButton.Values.DropDownArrowColor = Color.Empty;
            FixBankFormatsButton.Values.Image = Properties.Resources.paper;
            FixBankFormatsButton.Values.Text = "Ajustar Archivos Bancarios";
            FixBankFormatsButton.Click += AjustBankFormats_Click;
            // 
            // MovValidatorButton
            // 
            MovValidatorButton.Location = new Point(98, 189);
            MovValidatorButton.Name = "MovValidatorButton";
            MovValidatorButton.Size = new Size(285, 57);
            MovValidatorButton.StateCommon.Back.Color1 = Color.FromArgb(240, 244, 250);
            MovValidatorButton.StateCommon.Back.Color2 = Color.FromArgb(240, 244, 250);
            MovValidatorButton.StateCommon.Border.Color1 = Color.FromArgb(202, 215, 234);
            MovValidatorButton.StateCommon.Border.Rounding = 15F;
            MovValidatorButton.StateCommon.Border.Width = 1;
            MovValidatorButton.StateCommon.Content.ShortText.Color1 = Color.Black;
            MovValidatorButton.StateCommon.Content.ShortText.Font = new Font("Consolas", 10F, FontStyle.Bold);
            MovValidatorButton.StateTracking.Back.Color1 = Color.FromArgb(220, 230, 245);
            MovValidatorButton.StateTracking.Border.Color1 = Color.FromArgb(184, 204, 230);
            MovValidatorButton.TabIndex = 7;
            MovValidatorButton.TabStop = false;
            MovValidatorButton.Values.DropDownArrowColor = Color.Empty;
            MovValidatorButton.Values.Image = Properties.Resources.check;
            MovValidatorButton.Values.Text = "Validar Movimientos";
            MovValidatorButton.Click += MovValidator_Click;
            // 
            // FAQButton
            // 
            FAQButton.Location = new Point(98, 269);
            FAQButton.Name = "FAQButton";
            FAQButton.Size = new Size(285, 57);
            FAQButton.StateCommon.Back.Color1 = Color.FromArgb(240, 244, 250);
            FAQButton.StateCommon.Back.Color2 = Color.FromArgb(240, 244, 250);
            FAQButton.StateCommon.Border.Color1 = Color.FromArgb(202, 215, 234);
            FAQButton.StateCommon.Border.Rounding = 15F;
            FAQButton.StateCommon.Border.Width = 1;
            FAQButton.StateCommon.Content.ShortText.Color1 = Color.Black;
            FAQButton.StateCommon.Content.ShortText.Font = new Font("Consolas", 10F, FontStyle.Bold);
            FAQButton.StateTracking.Back.Color1 = Color.FromArgb(220, 230, 245);
            FAQButton.StateTracking.Border.Color1 = Color.FromArgb(184, 204, 230);
            FAQButton.TabIndex = 8;
            FAQButton.TabStop = false;
            FAQButton.Values.DropDownArrowColor = Color.Empty;
            FAQButton.Values.Image = Properties.Resources.help_web_button;
            FAQButton.Values.Text = "Saber más sobre ExcelFix";
            FAQButton.Click += FAQForm_Click;
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Font = new Font("Consolas", 6.75F, FontStyle.Regular, GraphicsUnit.Point, 0);
            label3.Location = new Point(3, 365);
            label3.Name = "label3";
            label3.Size = new Size(265, 10);
            label3.TabIndex = 9;
            label3.Text = "© 2025 ExcelFix · Desarrollado por Gabriel Rodriguez";
            label3.Click += label3_Click;
            // 
            // PrincipalMenu
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            BackColor = SystemColors.ActiveCaption;
            ClientSize = new Size(477, 381);
            Controls.Add(label3);
            Controls.Add(FAQButton);
            Controls.Add(MovValidatorButton);
            Controls.Add(FixBankFormatsButton);
            Controls.Add(label2);
            Controls.Add(label1);
            FormBorderStyle = FormBorderStyle.FixedSingle;
            Icon = (Icon)resources.GetObject("$this.Icon");
            Name = "PrincipalMenu";
            StartPosition = FormStartPosition.CenterScreen;
            Text = "ExcelFix";
            FormClosing += FormPrincipalMenu_FormClosing;
            Load += Form3_Load;
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Label label1;
        private Label label2;
        private Krypton.Toolkit.KryptonButton FixBankFormatsButton;
        private Krypton.Toolkit.KryptonButton MovValidatorButton;
        private Krypton.Toolkit.KryptonButton FAQButton;
        private Label label3;
    }
}
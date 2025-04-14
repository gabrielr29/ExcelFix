namespace Ajustador_de_formatos_Excel_de_movimientos_bancarios
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            executeButton = new Button();
            bankSelector = new ComboBox();
            tagBankSelector = new Label();
            button1 = new Button();
            tagFileDirection = new Label();
            textBoxPath = new TextBox();
            label1 = new Label();
            textBoxFileName = new TextBox();
            button2 = new Button();
            SuspendLayout();
            // 
            // executeButton
            // 
            executeButton.Font = new Font("Consolas", 12F, FontStyle.Bold, GraphicsUnit.Point, 0);
            executeButton.Location = new Point(145, 269);
            executeButton.Name = "executeButton";
            executeButton.Size = new Size(175, 48);
            executeButton.TabIndex = 0;
            executeButton.Text = "Ejecutar";
            executeButton.UseVisualStyleBackColor = true;
            executeButton.Click += Executebutton_Click;
            // 
            // bankSelector
            // 
            bankSelector.FormattingEnabled = true;
            bankSelector.Items.AddRange(new object[] { "Banesco (Modificar)", "Banco de Venezuela (Modificar)", "Mercantil (Modificar)", "Exterior (Modificar)", "Banesco (Ubicar duplicados)", "Banco de Vnzla/Exterior (Ubicar duplicados)" });
            bankSelector.Location = new Point(76, 84);
            bankSelector.Name = "bankSelector";
            bankSelector.Size = new Size(320, 23);
            bankSelector.TabIndex = 1;
            // 
            // tagBankSelector
            // 
            tagBankSelector.AutoSize = true;
            tagBankSelector.Font = new Font("Consolas", 12F, FontStyle.Regular, GraphicsUnit.Point, 0);
            tagBankSelector.Location = new Point(76, 62);
            tagBankSelector.Name = "tagBankSelector";
            tagBankSelector.Size = new Size(306, 19);
            tagBankSelector.TabIndex = 2;
            tagBankSelector.Text = "Tipo de formato Excel y operación";
            tagBankSelector.Click += label1_Click;
            // 
            // button1
            // 
            button1.Font = new Font("Consolas", 12F, FontStyle.Bold, GraphicsUnit.Point, 0);
            button1.Location = new Point(76, 119);
            button1.Name = "button1";
            button1.Size = new Size(98, 27);
            button1.TabIndex = 3;
            button1.Text = "Adjuntar banco";
            button1.UseVisualStyleBackColor = true;
            button1.Click += button1_Click_1;
            // 
            // tagFileDirection
            // 
            tagFileDirection.AutoSize = true;
            tagFileDirection.Font = new Font("Consolas", 10F);
            tagFileDirection.Location = new Point(76, 154);
            tagFileDirection.Name = "tagFileDirection";
            tagFileDirection.Size = new Size(152, 17);
            tagFileDirection.TabIndex = 4;
            tagFileDirection.Text = "Nombre Del Formato";
            // 
            // textBoxPath
            // 
            textBoxPath.Enabled = false;
            textBoxPath.Location = new Point(76, 174);
            textBoxPath.Name = "textBoxPath";
            textBoxPath.Size = new Size(320, 23);
            textBoxPath.TabIndex = 5;
            textBoxPath.TextChanged += textBox1_TextChanged;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Font = new Font("Consolas", 10F);
            label1.Location = new Point(76, 203);
            label1.Name = "label1";
            label1.Size = new Size(152, 17);
            label1.TabIndex = 6;
            label1.Text = "Nombre Del Archivo";
            // 
            // textBoxFileName
            // 
            textBoxFileName.Enabled = false;
            textBoxFileName.Location = new Point(76, 225);
            textBoxFileName.Name = "textBoxFileName";
            textBoxFileName.Size = new Size(320, 23);
            textBoxFileName.TabIndex = 7;
            textBoxFileName.TextChanged += textBoxFileName_TextChanged;
            // 
            // button2
            // 
            button2.Font = new Font("Consolas", 12F, FontStyle.Bold);
            button2.Location = new Point(12, 12);
            button2.Name = "button2";
            button2.Size = new Size(72, 29);
            button2.TabIndex = 14;
            button2.Text = "<";
            button2.UseVisualStyleBackColor = true;
            button2.Click += button2_Click;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            BackColor = SystemColors.ActiveCaption;
            ClientSize = new Size(467, 331);
            Controls.Add(button2);
            Controls.Add(textBoxFileName);
            Controls.Add(label1);
            Controls.Add(textBoxPath);
            Controls.Add(tagFileDirection);
            Controls.Add(button1);
            Controls.Add(tagBankSelector);
            Controls.Add(bankSelector);
            Controls.Add(executeButton);
            FormBorderStyle = FormBorderStyle.FixedSingle;
            Icon = (Icon)resources.GetObject("$this.Icon");
            MaximizeBox = false;
            Name = "Form1";
            StartPosition = FormStartPosition.CenterScreen;
            Text = "ExcelFix";
            FormClosing += FormBankFormatFixer_FormClosing;
            Load += Form1_Load;
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button executeButton;
        private ComboBox bankSelector;
        private Label tagBankSelector;
        private Button button1;
        private Label tagFileDirection;
        private TextBox textBoxPath;
        private Label label1;
        private TextBox textBoxFileName;
        private Button button2;
    }
}

namespace Ajustador_de_formatos_Excel_de_movimientos_bancarios
{
    partial class ExcelFixForm2
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ExcelFixForm2));
            dataGridView1 = new DataGridView();
            DateColumn = new DataGridViewTextBoxColumn();
            ValidationDate = new DataGridViewTextBoxColumn();
            ReferenceNumber = new DataGridViewTextBoxColumn();
            Description = new DataGridViewTextBoxColumn();
            Incomes = new DataGridViewTextBoxColumn();
            Expenses = new DataGridViewTextBoxColumn();
            BillNumber = new DataGridViewTextBoxColumn();
            CostumerCode = new DataGridViewTextBoxColumn();
            NRow = new DataGridViewTextBoxColumn();
            searchMovButton = new Button();
            bankMovMountTextBox = new TextBox();
            FileNameTextBox = new TextBox();
            dateTitleLabel = new Label();
            label2 = new Label();
            billCodeTextBox = new TextBox();
            billCodeLabel = new Label();
            updateRowButton = new Button();
            clientCodeTextBox = new TextBox();
            label1 = new Label();
            button1 = new Button();
            ReferenceNumberTextBox = new TextBox();
            button2 = new Button();
            label3 = new Label();
            createCopy = new Button();
            ((System.ComponentModel.ISupportInitialize)dataGridView1).BeginInit();
            SuspendLayout();
            // 
            // dataGridView1
            // 
            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.AllowUserToDeleteRows = false;
            dataGridView1.AllowUserToResizeRows = false;
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridView1.Columns.AddRange(new DataGridViewColumn[] { DateColumn, ValidationDate, ReferenceNumber, Description, Incomes, Expenses, BillNumber, CostumerCode, NRow });
            dataGridView1.Location = new Point(29, 194);
            dataGridView1.MultiSelect = false;
            dataGridView1.Name = "dataGridView1";
            dataGridView1.ReadOnly = true;
            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridView1.Size = new Size(1275, 227);
            dataGridView1.TabIndex = 0;
            dataGridView1.CellContentClick += dataGridView1_CellContentClick;
            // 
            // DateColumn
            // 
            DateColumn.FillWeight = 78.23603F;
            DateColumn.HeaderText = "Fecha";
            DateColumn.Name = "DateColumn";
            DateColumn.ReadOnly = true;
            // 
            // ValidationDate
            // 
            ValidationDate.FillWeight = 78.23603F;
            ValidationDate.HeaderText = "Fecha Validación";
            ValidationDate.Name = "ValidationDate";
            ValidationDate.ReadOnly = true;
            // 
            // ReferenceNumber
            // 
            ReferenceNumber.FillWeight = 78.23603F;
            ReferenceNumber.HeaderText = "Referencia";
            ReferenceNumber.Name = "ReferenceNumber";
            ReferenceNumber.ReadOnly = true;
            // 
            // Description
            // 
            Description.FillWeight = 274.111664F;
            Description.HeaderText = "Descripción";
            Description.Name = "Description";
            Description.ReadOnly = true;
            // 
            // Incomes
            // 
            Incomes.FillWeight = 78.23603F;
            Incomes.HeaderText = "Ingresos";
            Incomes.Name = "Incomes";
            Incomes.ReadOnly = true;
            // 
            // Expenses
            // 
            Expenses.FillWeight = 78.23603F;
            Expenses.HeaderText = "Egresos";
            Expenses.Name = "Expenses";
            Expenses.ReadOnly = true;
            // 
            // BillNumber
            // 
            BillNumber.FillWeight = 78.23603F;
            BillNumber.HeaderText = "Número de factura";
            BillNumber.Name = "BillNumber";
            BillNumber.ReadOnly = true;
            // 
            // CostumerCode
            // 
            CostumerCode.FillWeight = 78.23603F;
            CostumerCode.HeaderText = "Codigo de cliente";
            CostumerCode.Name = "CostumerCode";
            CostumerCode.ReadOnly = true;
            // 
            // NRow
            // 
            NRow.FillWeight = 78.23603F;
            NRow.HeaderText = "N° Fila";
            NRow.Name = "NRow";
            NRow.ReadOnly = true;
            // 
            // searchMovButton
            // 
            searchMovButton.Font = new Font("Consolas", 14.25F, FontStyle.Bold);
            searchMovButton.Location = new Point(849, 141);
            searchMovButton.Name = "searchMovButton";
            searchMovButton.Size = new Size(190, 36);
            searchMovButton.TabIndex = 3;
            searchMovButton.Text = "Buscar";
            searchMovButton.UseVisualStyleBackColor = true;
            searchMovButton.Click += SearchMoveButton_Click;
            // 
            // bankMovMountTextBox
            // 
            bankMovMountTextBox.Font = new Font("Consolas", 12F);
            bankMovMountTextBox.Location = new Point(325, 150);
            bankMovMountTextBox.MaxLength = 200;
            bankMovMountTextBox.Name = "bankMovMountTextBox";
            bankMovMountTextBox.Size = new Size(221, 26);
            bankMovMountTextBox.TabIndex = 1;
            bankMovMountTextBox.TextChanged += bankMovMountTextBox_TextChanged;
            bankMovMountTextBox.KeyPress += bankMovMountTextBox_KeyPress;
            // 
            // FileNameTextBox
            // 
            FileNameTextBox.Font = new Font("Consolas", 12F);
            FileNameTextBox.Location = new Point(578, 79);
            FileNameTextBox.Name = "FileNameTextBox";
            FileNameTextBox.ReadOnly = true;
            FileNameTextBox.Size = new Size(250, 26);
            FileNameTextBox.TabIndex = 55;
            FileNameTextBox.TabStop = false;
            FileNameTextBox.TextChanged += ReferenceNumberTextBox_TextChanged;
            // 
            // dateTitleLabel
            // 
            dateTitleLabel.AutoSize = true;
            dateTitleLabel.Font = new Font("Consolas", 12F);
            dateTitleLabel.Location = new Point(325, 127);
            dateTitleLabel.Name = "dateTitleLabel";
            dateTitleLabel.Size = new Size(54, 19);
            dateTitleLabel.TabIndex = 15;
            dateTitleLabel.Text = "Monto";
            dateTitleLabel.Click += dateTitleLabel_Click;
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Font = new Font("Consolas", 12F);
            label2.Location = new Point(578, 127);
            label2.Name = "label2";
            label2.Size = new Size(153, 19);
            label2.TabIndex = 17;
            label2.Text = "N# de referencia";
            // 
            // billCodeTextBox
            // 
            billCodeTextBox.Font = new Font("Consolas", 12F);
            billCodeTextBox.Location = new Point(372, 474);
            billCodeTextBox.Name = "billCodeTextBox";
            billCodeTextBox.Size = new Size(280, 26);
            billCodeTextBox.TabIndex = 4;
            billCodeTextBox.TextChanged += billCodeTextBox_TextChanged;
            billCodeTextBox.KeyPress += billCodeTextBox_KeyPress;
            // 
            // billCodeLabel
            // 
            billCodeLabel.AutoSize = true;
            billCodeLabel.Font = new Font("Consolas", 12F);
            billCodeLabel.Location = new Point(372, 452);
            billCodeLabel.Name = "billCodeLabel";
            billCodeLabel.Size = new Size(135, 19);
            billCodeLabel.TabIndex = 9;
            billCodeLabel.Text = "N° de facturas";
            // 
            // updateRowButton
            // 
            updateRowButton.Font = new Font("Consolas", 14.25F, FontStyle.Bold);
            updateRowButton.Location = new Point(523, 519);
            updateRowButton.Name = "updateRowButton";
            updateRowButton.Size = new Size(276, 49);
            updateRowButton.TabIndex = 6;
            updateRowButton.Text = "Actualizar ";
            updateRowButton.UseVisualStyleBackColor = true;
            updateRowButton.Click += updateRowButton_Click;
            // 
            // clientCodeTextBox
            // 
            clientCodeTextBox.Font = new Font("Consolas", 12F);
            clientCodeTextBox.Location = new Point(681, 474);
            clientCodeTextBox.Name = "clientCodeTextBox";
            clientCodeTextBox.Size = new Size(310, 26);
            clientCodeTextBox.TabIndex = 5;
            clientCodeTextBox.TextChanged += clientCodeTextBox_TextChanged;
            clientCodeTextBox.KeyPress += clientCodeTextBox_KeyPress;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Font = new Font("Consolas", 12F);
            label1.Location = new Point(681, 452);
            label1.Name = "label1";
            label1.Size = new Size(171, 19);
            label1.TabIndex = 12;
            label1.Text = "Código del cliente";
            // 
            // button1
            // 
            button1.Font = new Font("Consolas", 12F, FontStyle.Bold);
            button1.Location = new Point(29, 23);
            button1.Name = "button1";
            button1.Size = new Size(83, 29);
            button1.TabIndex = 13;
            button1.TabStop = false;
            button1.Text = "<";
            button1.UseVisualStyleBackColor = true;
            button1.Click += button1_Click;
            // 
            // ReferenceNumberTextBox
            // 
            ReferenceNumberTextBox.Font = new Font("Consolas", 12F);
            ReferenceNumberTextBox.Location = new Point(581, 150);
            ReferenceNumberTextBox.MaxLength = 200;
            ReferenceNumberTextBox.Name = "ReferenceNumberTextBox";
            ReferenceNumberTextBox.Size = new Size(247, 26);
            ReferenceNumberTextBox.TabIndex = 2;
            ReferenceNumberTextBox.TextChanged += ReferenceNumberTextBox_TextChanged_1;
            ReferenceNumberTextBox.KeyPress += ReferenceNumberTextBox_KeyPress;
            // 
            // button2
            // 
            button2.Font = new Font("Consolas", 14.25F, FontStyle.Bold);
            button2.Location = new Point(325, 56);
            button2.Name = "button2";
            button2.Size = new Size(221, 49);
            button2.TabIndex = 15;
            button2.Text = "Adjuntar";
            button2.UseVisualStyleBackColor = true;
            button2.Click += attachFilebutton_Click;
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Font = new Font("Consolas", 12F);
            label3.Location = new Point(578, 59);
            label3.Name = "label3";
            label3.Size = new Size(243, 19);
            label3.TabIndex = 19;
            label3.Text = "Nombre del archivo (Excel)";
            // 
            // createCopy
            // 
            createCopy.Font = new Font("Consolas", 12F, FontStyle.Bold, GraphicsUnit.Point, 0);
            createCopy.Location = new Point(29, 128);
            createCopy.Name = "createCopy";
            createCopy.Size = new Size(190, 49);
            createCopy.TabIndex = 56;
            createCopy.Text = "Crear Copia de Seguridad";
            createCopy.UseVisualStyleBackColor = true;
            createCopy.Click += createCopy_Click;
            // 
            // ExcelFixForm2
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            BackColor = SystemColors.ActiveCaption;
            BackgroundImageLayout = ImageLayout.None;
            ClientSize = new Size(1349, 652);
            Controls.Add(createCopy);
            Controls.Add(label3);
            Controls.Add(button2);
            Controls.Add(ReferenceNumberTextBox);
            Controls.Add(button1);
            Controls.Add(label1);
            Controls.Add(clientCodeTextBox);
            Controls.Add(updateRowButton);
            Controls.Add(billCodeLabel);
            Controls.Add(billCodeTextBox);
            Controls.Add(label2);
            Controls.Add(dateTitleLabel);
            Controls.Add(FileNameTextBox);
            Controls.Add(bankMovMountTextBox);
            Controls.Add(searchMovButton);
            Controls.Add(dataGridView1);
            FormBorderStyle = FormBorderStyle.FixedSingle;
            Icon = (Icon)resources.GetObject("$this.Icon");
            Name = "ExcelFixForm2";
            StartPosition = FormStartPosition.CenterScreen;
            Tag = "";
            Text = "Validador de Movimientos";
            WindowState = FormWindowState.Maximized;
            FormClosing += FormMovValidator_FormClosing;
            Load += ExcelFixForm2_Load;
            ((System.ComponentModel.ISupportInitialize)dataGridView1).EndInit();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private DataGridView dataGridView1;
        private Button searchMovButton;
        private TextBox bankMovMountTextBox;
        private TextBox FileNameTextBox;
        private Label dateTitleLabel;
        private Label label2;
        private TextBox billCodeTextBox;
        private Label billCodeLabel;
        private Button updateRowButton;
        private TextBox clientCodeTextBox;
        private Label label1;
        private Button button1;
        private TextBox ReferenceNumberTextBox;
        private Button button2;
        private Label label3;
        private Button createCopy;
        private DataGridViewTextBoxColumn DateColumn;
        private DataGridViewTextBoxColumn ValidationDate;
        private DataGridViewTextBoxColumn ReferenceNumber;
        private DataGridViewTextBoxColumn Description;
        private DataGridViewTextBoxColumn Incomes;
        private DataGridViewTextBoxColumn Expenses;
        private DataGridViewTextBoxColumn BillNumber;
        private DataGridViewTextBoxColumn CostumerCode;
        private DataGridViewTextBoxColumn NRow;
    }
}
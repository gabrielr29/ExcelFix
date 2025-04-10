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
            dataGridView1.Location = new Point(12, 152);
            dataGridView1.MultiSelect = false;
            dataGridView1.Name = "dataGridView1";
            dataGridView1.ReadOnly = true;
            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridView1.Size = new Size(860, 150);
            dataGridView1.TabIndex = 0;
            dataGridView1.CellContentClick += dataGridView1_CellContentClick;
            // 
            // DateColumn
            // 
            DateColumn.HeaderText = "Fecha";
            DateColumn.Name = "DateColumn";
            DateColumn.ReadOnly = true;
            // 
            // ValidationDate
            // 
            ValidationDate.HeaderText = "Fecha Validación";
            ValidationDate.Name = "ValidationDate";
            ValidationDate.ReadOnly = true;
            // 
            // ReferenceNumber
            // 
            ReferenceNumber.HeaderText = "Referencia";
            ReferenceNumber.Name = "ReferenceNumber";
            ReferenceNumber.ReadOnly = true;
            // 
            // Description
            // 
            Description.HeaderText = "Descripción";
            Description.Name = "Description";
            Description.ReadOnly = true;
            // 
            // Incomes
            // 
            Incomes.HeaderText = "Ingresos";
            Incomes.Name = "Incomes";
            Incomes.ReadOnly = true;
            // 
            // Expenses
            // 
            Expenses.HeaderText = "Egresos";
            Expenses.Name = "Expenses";
            Expenses.ReadOnly = true;
            // 
            // BillNumber
            // 
            BillNumber.HeaderText = "Número de factura";
            BillNumber.Name = "BillNumber";
            BillNumber.ReadOnly = true;
            // 
            // CostumerCode
            // 
            CostumerCode.HeaderText = "Codigo de cliente";
            CostumerCode.Name = "CostumerCode";
            CostumerCode.ReadOnly = true;
            // 
            // NRow
            // 
            NRow.HeaderText = "N° Fila";
            NRow.Name = "NRow";
            NRow.ReadOnly = true;
            // 
            // searchMovButton
            // 
            searchMovButton.Font = new Font("Consolas", 12F, FontStyle.Bold);
            searchMovButton.Location = new Point(577, 100);
            searchMovButton.Name = "searchMovButton";
            searchMovButton.Size = new Size(190, 36);
            searchMovButton.TabIndex = 3;
            searchMovButton.Text = "Buscar";
            searchMovButton.UseVisualStyleBackColor = true;
            searchMovButton.Click += SearchMoveButton_Click;
            // 
            // bankMovMountTextBox
            // 
            bankMovMountTextBox.Location = new Point(53, 109);
            bankMovMountTextBox.Name = "bankMovMountTextBox";
            bankMovMountTextBox.Size = new Size(221, 23);
            bankMovMountTextBox.TabIndex = 1;
            // 
            // FileNameTextBox
            // 
            FileNameTextBox.Location = new Point(306, 38);
            FileNameTextBox.Name = "FileNameTextBox";
            FileNameTextBox.ReadOnly = true;
            FileNameTextBox.Size = new Size(250, 23);
            FileNameTextBox.TabIndex = 55;
            FileNameTextBox.TabStop = false;
            FileNameTextBox.TextChanged += ReferenceNumberTextBox_TextChanged;
            // 
            // dateTitleLabel
            // 
            dateTitleLabel.AutoSize = true;
            dateTitleLabel.Font = new Font("Consolas", 10F);
            dateTitleLabel.Location = new Point(53, 86);
            dateTitleLabel.Name = "dateTitleLabel";
            dateTitleLabel.Size = new Size(48, 17);
            dateTitleLabel.TabIndex = 15;
            dateTitleLabel.Text = "Monto";
            dateTitleLabel.Click += dateTitleLabel_Click;
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Font = new Font("Consolas", 10F);
            label2.Location = new Point(306, 86);
            label2.Name = "label2";
            label2.Size = new Size(136, 17);
            label2.TabIndex = 17;
            label2.Text = "N# de referencia";
            // 
            // billCodeTextBox
            // 
            billCodeTextBox.Location = new Point(154, 345);
            billCodeTextBox.Name = "billCodeTextBox";
            billCodeTextBox.Size = new Size(280, 23);
            billCodeTextBox.TabIndex = 6;
            // 
            // billCodeLabel
            // 
            billCodeLabel.AutoSize = true;
            billCodeLabel.Font = new Font("Consolas", 10F);
            billCodeLabel.Location = new Point(153, 325);
            billCodeLabel.Name = "billCodeLabel";
            billCodeLabel.Size = new Size(120, 17);
            billCodeLabel.TabIndex = 9;
            billCodeLabel.Text = "N° de facturas";
            // 
            // updateRowButton
            // 
            updateRowButton.Font = new Font("Consolas", 12F, FontStyle.Bold);
            updateRowButton.Location = new Point(297, 389);
            updateRowButton.Name = "updateRowButton";
            updateRowButton.Size = new Size(244, 49);
            updateRowButton.TabIndex = 7;
            updateRowButton.Text = "Actualizar ";
            updateRowButton.UseVisualStyleBackColor = true;
            updateRowButton.Click += updateRowButton_Click;
            // 
            // clientCodeTextBox
            // 
            clientCodeTextBox.Location = new Point(464, 345);
            clientCodeTextBox.Name = "clientCodeTextBox";
            clientCodeTextBox.Size = new Size(250, 23);
            clientCodeTextBox.TabIndex = 5;
            clientCodeTextBox.TextChanged += clientCodeTextBox_TextChanged;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Font = new Font("Consolas", 10F);
            label1.Location = new Point(464, 325);
            label1.Name = "label1";
            label1.Size = new Size(152, 17);
            label1.TabIndex = 12;
            label1.Text = "Código del cliente";
            // 
            // button1
            // 
            button1.Font = new Font("Consolas", 12F, FontStyle.Bold);
            button1.Location = new Point(12, 12);
            button1.Name = "button1";
            button1.Size = new Size(72, 29);
            button1.TabIndex = 13;
            button1.TabStop = false;
            button1.Text = "<";
            button1.UseVisualStyleBackColor = true;
            button1.Click += button1_Click;
            // 
            // ReferenceNumberTextBox
            // 
            ReferenceNumberTextBox.Location = new Point(309, 109);
            ReferenceNumberTextBox.Name = "ReferenceNumberTextBox";
            ReferenceNumberTextBox.Size = new Size(247, 23);
            ReferenceNumberTextBox.TabIndex = 2;
            ReferenceNumberTextBox.TextChanged += ReferenceNumberTextBox_TextChanged_1;
            ReferenceNumberTextBox.KeyPress += ReferenceNumberTextBox_KeyPress;
            // 
            // button2
            // 
            button2.Font = new Font("Consolas", 12F, FontStyle.Bold);
            button2.Location = new Point(109, 12);
            button2.Name = "button2";
            button2.Size = new Size(165, 49);
            button2.TabIndex = 15;
            button2.Text = "Adjuntar";
            button2.UseVisualStyleBackColor = true;
            button2.Click += button2_Click;
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Font = new Font("Consolas", 10F);
            label3.Location = new Point(306, 18);
            label3.Name = "label3";
            label3.Size = new Size(216, 17);
            label3.TabIndex = 19;
            label3.Text = "Nombre del archivo (Excel)";
            // 
            // ExcelFixForm2
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            BackColor = SystemColors.ActiveCaption;
            BackgroundImageLayout = ImageLayout.None;
            ClientSize = new Size(884, 450);
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
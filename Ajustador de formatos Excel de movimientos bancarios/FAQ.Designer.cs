using System.Windows.Forms;

namespace Ajustador_de_formatos_Excel_de_movimientos_bancarios
{
    partial class FAQ
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FAQ));
            FAQRichTextBox = new Krypton.Toolkit.KryptonRichTextBox();
            TitleLabel = new Label();
            button1 = new Button();
            SuspendLayout();
            // 
            // FAQRichTextBox
            // 
            FAQRichTextBox.Location = new Point(12, 66);
            FAQRichTextBox.Name = "FAQRichTextBox";
            FAQRichTextBox.ReadOnly = true;
            FAQRichTextBox.ScrollBars = RichTextBoxScrollBars.Vertical;
            FAQRichTextBox.Size = new Size(453, 303);
            FAQRichTextBox.TabIndex = 0;
            FAQRichTextBox.Text = resources.GetString("FAQRichTextBox.Text");
            FAQRichTextBox.TextChanged += kryptonRichTextBox1_TextChanged;
            // 
            // TitleLabel
            // 
            TitleLabel.AutoSize = true;
            TitleLabel.Font = new Font("Consolas", 12F, FontStyle.Bold, GraphicsUnit.Point, 0);
            TitleLabel.ForeColor = Color.DarkBlue;
            TitleLabel.Location = new Point(142, 21);
            TitleLabel.Name = "TitleLabel";
            TitleLabel.Size = new Size(171, 19);
            TitleLabel.TabIndex = 1;
            TitleLabel.Text = "Sobre Excel Fix...";
            TitleLabel.Click += label1_Click;
            // 
            // button1
            // 
            button1.Font = new Font("Consolas", 12F, FontStyle.Bold);
            button1.Location = new Point(11, 16);
            button1.Name = "button1";
            button1.Size = new Size(72, 29);
            button1.TabIndex = 14;
            button1.TabStop = false;
            button1.Text = "<";
            button1.UseVisualStyleBackColor = true;
            button1.Click += button1_Click;
            // 
            // FAQ
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            BackColor = SystemColors.ActiveCaption;
            ClientSize = new Size(477, 381);
            Controls.Add(button1);
            Controls.Add(TitleLabel);
            Controls.Add(FAQRichTextBox);
            FormBorderStyle = FormBorderStyle.FixedSingle;
            Icon = (Icon)resources.GetObject("$this.Icon");
            Name = "FAQ";
            StartPosition = FormStartPosition.CenterScreen;
            Text = "FAQ";
            FormClosing += FormFAQ_FormClosing;
            Load += FAQ_Load;
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Krypton.Toolkit.KryptonRichTextBox FAQRichTextBox;
        private Label TitleLabel;
        private Button button1;
    }
}
namespace Kwerendy
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
            btnLoad = new Button();
            txtBoxTest = new TextBox();
            grid1 = new DataGridView();
            btnRegistrations = new Button();
            butGenerate = new Button();
            btnSignature = new Button();
            ((System.ComponentModel.ISupportInitialize)grid1).BeginInit();
            SuspendLayout();
            // 
            // btnLoad
            // 
            btnLoad.Font = new Font("Segoe UI", 20F);
            btnLoad.Location = new Point(47, 31);
            btnLoad.Name = "btnLoad";
            btnLoad.Size = new Size(197, 47);
            btnLoad.TabIndex = 0;
            btnLoad.Text = "Wybierz plik";
            btnLoad.UseVisualStyleBackColor = true;
            btnLoad.Click += btnLoad_Click;
            // 
            // txtBoxTest
            // 
            txtBoxTest.Location = new Point(47, 540);
            txtBoxTest.Name = "txtBoxTest";
            txtBoxTest.Size = new Size(197, 23);
            txtBoxTest.TabIndex = 1;
            // 
            // grid1
            // 
            grid1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            grid1.Location = new Point(442, 31);
            grid1.Name = "grid1";
            grid1.Size = new Size(804, 532);
            grid1.TabIndex = 2;
            // 
            // btnRegistrations
            // 
            btnRegistrations.Font = new Font("Segoe UI", 15F);
            btnRegistrations.Location = new Point(47, 84);
            btnRegistrations.Name = "btnRegistrations";
            btnRegistrations.Size = new Size(197, 47);
            btnRegistrations.TabIndex = 3;
            btnRegistrations.Text = "Tylko rejestracje";
            btnRegistrations.UseVisualStyleBackColor = true;
            btnRegistrations.Click += btnRegistrations_Click;
            // 
            // butGenerate
            // 
            butGenerate.Font = new Font("Segoe UI", 15F);
            butGenerate.Location = new Point(47, 190);
            butGenerate.Name = "butGenerate";
            butGenerate.Size = new Size(197, 47);
            butGenerate.TabIndex = 4;
            butGenerate.Text = "Generuj";
            butGenerate.UseVisualStyleBackColor = true;
            butGenerate.Click += butGenerate_Click;
            // 
            // btnSignature
            // 
            btnSignature.Font = new Font("Segoe UI", 15F);
            btnSignature.Location = new Point(47, 137);
            btnSignature.Name = "btnSignature";
            btnSignature.Size = new Size(197, 47);
            btnSignature.TabIndex = 5;
            btnSignature.Text = "Podpis";
            btnSignature.UseVisualStyleBackColor = true;
            btnSignature.Click += btnSignature_Click;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            AutoValidate = AutoValidate.EnablePreventFocusChange;
            ClientSize = new Size(1316, 649);
            Controls.Add(btnSignature);
            Controls.Add(butGenerate);
            Controls.Add(btnRegistrations);
            Controls.Add(grid1);
            Controls.Add(txtBoxTest);
            Controls.Add(btnLoad);
            Name = "Form1";
            Text = "Form1";
            ((System.ComponentModel.ISupportInitialize)grid1).EndInit();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button btnLoad;
        private TextBox txtBoxTest;
        private DataGridView grid1;
        private Button btnRegistrations;
        private Button butGenerate;
        private Button btnSignature;
    }
}

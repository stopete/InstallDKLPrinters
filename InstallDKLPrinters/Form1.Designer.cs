namespace InstallDKLPrinters
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
            components = new System.ComponentModel.Container();
            btnInstallPrinters = new Button();
            textBox1 = new TextBox();
            progressBar1 = new ProgressBar();
            lblProgress = new Label();
            timer1 = new System.Windows.Forms.Timer(components);
            SuspendLayout();
            // 
            // btnInstallPrinters
            // 
            btnInstallPrinters.Location = new Point(91, 31);
            btnInstallPrinters.Name = "btnInstallPrinters";
            btnInstallPrinters.Size = new Size(161, 29);
            btnInstallPrinters.TabIndex = 0;
            btnInstallPrinters.Text = "Install DKL Printers";
            btnInstallPrinters.UseVisualStyleBackColor = true;
            btnInstallPrinters.Click += btnInstallPrinters_Click;
            // 
            // textBox1
            // 
            textBox1.Location = new Point(85, 80);
            textBox1.Multiline = true;
            textBox1.Name = "textBox1";
            textBox1.ScrollBars = ScrollBars.Vertical;
            textBox1.Size = new Size(504, 223);
            textBox1.TabIndex = 1;
            // 
            // progressBar1
            // 
            progressBar1.Location = new Point(85, 344);
            progressBar1.Name = "progressBar1";
            progressBar1.Size = new Size(494, 29);
            progressBar1.TabIndex = 2;
            // 
            // lblProgress
            // 
            lblProgress.AutoSize = true;
            lblProgress.Location = new Point(91, 318);
            lblProgress.Name = "lblProgress";
            lblProgress.Size = new Size(29, 20);
            lblProgress.TabIndex = 3;
            lblProgress.Text = "0%";
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(8F, 20F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(669, 440);
            Controls.Add(lblProgress);
            Controls.Add(progressBar1);
            Controls.Add(textBox1);
            Controls.Add(btnInstallPrinters);
            Name = "Form1";
            Text = "Install DKL Printers";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button btnInstallPrinters;
        private TextBox textBox1;
        private ProgressBar progressBar1;
        private Label lblProgress;
        private System.Windows.Forms.Timer timer1;
    }
}

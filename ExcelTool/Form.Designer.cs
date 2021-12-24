namespace ExcelTool
{
    partial class Form
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
            this.buttonInputDirectory = new System.Windows.Forms.Button();
            this.listBoxMessage = new System.Windows.Forms.ListBox();
            this.backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
            this.buttonOutputDirectory = new System.Windows.Forms.Button();
            this.buttonExecute = new System.Windows.Forms.Button();
            this.textBoxInputDirectory = new System.Windows.Forms.TextBox();
            this.textBoxOutputDirectory = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // buttonInputDirectory
            // 
            this.buttonInputDirectory.Location = new System.Drawing.Point(713, 12);
            this.buttonInputDirectory.Name = "buttonInputDirectory";
            this.buttonInputDirectory.Size = new System.Drawing.Size(75, 23);
            this.buttonInputDirectory.TabIndex = 0;
            this.buttonInputDirectory.Text = "원본 경로";
            this.buttonInputDirectory.UseVisualStyleBackColor = true;
            this.buttonInputDirectory.Click += new System.EventHandler(this.buttonInputDirectory_Click);
            // 
            // listBoxMessage
            // 
            this.listBoxMessage.FormattingEnabled = true;
            this.listBoxMessage.ItemHeight = 15;
            this.listBoxMessage.Location = new System.Drawing.Point(12, 74);
            this.listBoxMessage.Name = "listBoxMessage";
            this.listBoxMessage.Size = new System.Drawing.Size(776, 364);
            this.listBoxMessage.TabIndex = 1;
            // 
            // buttonOutputDirectory
            // 
            this.buttonOutputDirectory.Location = new System.Drawing.Point(713, 41);
            this.buttonOutputDirectory.Name = "buttonOutputDirectory";
            this.buttonOutputDirectory.Size = new System.Drawing.Size(75, 23);
            this.buttonOutputDirectory.TabIndex = 2;
            this.buttonOutputDirectory.Text = "출력 경로";
            this.buttonOutputDirectory.UseVisualStyleBackColor = true;
            this.buttonOutputDirectory.Click += new System.EventHandler(this.buttonOutputDirectory_Click);
            // 
            // buttonExecute
            // 
            this.buttonExecute.Location = new System.Drawing.Point(369, 444);
            this.buttonExecute.Name = "buttonExecute";
            this.buttonExecute.Size = new System.Drawing.Size(75, 23);
            this.buttonExecute.TabIndex = 3;
            this.buttonExecute.Text = "실행";
            this.buttonExecute.UseVisualStyleBackColor = true;
            this.buttonExecute.Click += new System.EventHandler(this.buttonExecute_Click);
            // 
            // textBoxInputDirectory
            // 
            this.textBoxInputDirectory.Location = new System.Drawing.Point(12, 16);
            this.textBoxInputDirectory.Name = "textBoxInputDirectory";
            this.textBoxInputDirectory.Size = new System.Drawing.Size(695, 23);
            this.textBoxInputDirectory.TabIndex = 4;
            // 
            // textBoxOutputDirectory
            // 
            this.textBoxOutputDirectory.Location = new System.Drawing.Point(12, 45);
            this.textBoxOutputDirectory.Name = "textBoxOutputDirectory";
            this.textBoxOutputDirectory.Size = new System.Drawing.Size(695, 23);
            this.textBoxOutputDirectory.TabIndex = 5;
            // 
            // Form
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 474);
            this.Controls.Add(this.textBoxOutputDirectory);
            this.Controls.Add(this.textBoxInputDirectory);
            this.Controls.Add(this.buttonExecute);
            this.Controls.Add(this.buttonOutputDirectory);
            this.Controls.Add(this.listBoxMessage);
            this.Controls.Add(this.buttonInputDirectory);
            this.Name = "Form";
            this.Text = "ExcelTool";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private Button buttonInputDirectory;
        private ListBox listBoxMessage;
        private System.ComponentModel.BackgroundWorker backgroundWorker1;
        private Button buttonOutputDirectory;
        private Button buttonExecute;
        private TextBox textBoxInputDirectory;
        private TextBox textBoxOutputDirectory;
    }
}
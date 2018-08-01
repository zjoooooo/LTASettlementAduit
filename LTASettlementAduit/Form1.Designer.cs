namespace LTASettlementAduit
{
    partial class Form1
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
            System.Windows.Forms.Button ReadPMS;
            this.CompareServer = new System.Windows.Forms.Button();
            ReadPMS = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // ReadPMS
            // 
            ReadPMS.Location = new System.Drawing.Point(243, 189);
            ReadPMS.Name = "ReadPMS";
            ReadPMS.Size = new System.Drawing.Size(81, 31);
            ReadPMS.TabIndex = 0;
            ReadPMS.Text = "ReadPMS";
            ReadPMS.UseVisualStyleBackColor = true;
            ReadPMS.Click += new System.EventHandler(this.ReadPMS_Click);
            // 
            // CompareServer
            // 
            this.CompareServer.Location = new System.Drawing.Point(420, 189);
            this.CompareServer.Name = "CompareServer";
            this.CompareServer.Size = new System.Drawing.Size(126, 31);
            this.CompareServer.TabIndex = 1;
            this.CompareServer.Text = "CompareWithServer";
            this.CompareServer.UseVisualStyleBackColor = true;
            this.CompareServer.Click += new System.EventHandler(this.CompareServer_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.CompareServer);
            this.Controls.Add(ReadPMS);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button CompareServer;
    }
}


namespace WordAddInResizeImage
{
    partial class FloatingPanel
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
            this.lkBtn = new System.Windows.Forms.LinkLabel();
            this.tbTooltip = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // lkBtn
            // 
            this.lkBtn.AutoSize = true;
            this.lkBtn.Location = new System.Drawing.Point(211, 110);
            this.lkBtn.Name = "lkBtn";
            this.lkBtn.Size = new System.Drawing.Size(47, 13);
            this.lkBtn.TabIndex = 4;
            this.lkBtn.TabStop = true;
            this.lkBtn.Text = "Remove";
            this.lkBtn.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lkBtn_LinkClicked);
            // 
            // tbTooltip
            // 
            this.tbTooltip.Font = new System.Drawing.Font("SimSun", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.tbTooltip.Location = new System.Drawing.Point(1, 12);
            this.tbTooltip.Multiline = true;
            this.tbTooltip.Name = "tbTooltip";
            this.tbTooltip.ReadOnly = true;
            this.tbTooltip.Size = new System.Drawing.Size(257, 92);
            this.tbTooltip.TabIndex = 5;
            // 
            // FloatingPanel
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(281, 132);
            this.Controls.Add(this.tbTooltip);
            this.Controls.Add(this.lkBtn);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "FloatingPanel";
            this.Text = "FloatingPanel";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.LinkLabel lkBtn;
        private System.Windows.Forms.TextBox tbTooltip;
    }
}
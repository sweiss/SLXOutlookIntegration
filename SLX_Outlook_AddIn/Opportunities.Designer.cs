namespace SLX_Outlook_AddIn
{
    partial class Opportunities
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
            this.grdOpportunities = new System.Windows.Forms.DataGridView();
            this.grdCmdDetails = new System.Windows.Forms.DataGridViewButtonColumn();
            this.grdCmdOpen = new System.Windows.Forms.DataGridViewButtonColumn();
            this.txtOppoDesc = new System.Windows.Forms.TextBox();
            ((System.ComponentModel.ISupportInitialize)(this.grdOpportunities)).BeginInit();
            this.SuspendLayout();
            // 
            // grdOpportunities
            // 
            this.grdOpportunities.AllowUserToAddRows = false;
            this.grdOpportunities.AllowUserToDeleteRows = false;
            this.grdOpportunities.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.grdOpportunities.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.grdCmdDetails,
            this.grdCmdOpen});
            this.grdOpportunities.Dock = System.Windows.Forms.DockStyle.Top;
            this.grdOpportunities.Location = new System.Drawing.Point(0, 0);
            this.grdOpportunities.MultiSelect = false;
            this.grdOpportunities.Name = "grdOpportunities";
            this.grdOpportunities.ReadOnly = true;
            this.grdOpportunities.Size = new System.Drawing.Size(717, 171);
            this.grdOpportunities.TabIndex = 0;
            this.grdOpportunities.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.grdOpportunities_CellContentClick);
            // 
            // grdCmdDetails
            // 
            this.grdCmdDetails.HeaderText = "";
            this.grdCmdDetails.Name = "grdCmdDetails";
            this.grdCmdDetails.ReadOnly = true;
            this.grdCmdDetails.Text = "Details";
            this.grdCmdDetails.UseColumnTextForButtonValue = true;
            this.grdCmdDetails.Width = 50;
            // 
            // grdCmdOpen
            // 
            this.grdCmdOpen.HeaderText = "";
            this.grdCmdOpen.Name = "grdCmdOpen";
            this.grdCmdOpen.ReadOnly = true;
            this.grdCmdOpen.Text = "Open";
            this.grdCmdOpen.UseColumnTextForButtonValue = true;
            this.grdCmdOpen.Width = 50;
            // 
            // txtOppoDesc
            // 
            this.txtOppoDesc.Location = new System.Drawing.Point(12, 177);
            this.txtOppoDesc.Name = "txtOppoDesc";
            this.txtOppoDesc.Size = new System.Drawing.Size(100, 20);
            this.txtOppoDesc.TabIndex = 1;
            // 
            // Opportunities
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(717, 382);
            this.Controls.Add(this.txtOppoDesc);
            this.Controls.Add(this.grdOpportunities);
            this.Name = "Opportunities";
            this.Text = "Opportunities";
            this.Load += new System.EventHandler(this.Opportunities_Load);
            ((System.ComponentModel.ISupportInitialize)(this.grdOpportunities)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView grdOpportunities;
        private System.Windows.Forms.DataGridViewButtonColumn grdCmdDetails;
        private System.Windows.Forms.DataGridViewButtonColumn grdCmdOpen;
        private System.Windows.Forms.TextBox txtOppoDesc;
    }
}
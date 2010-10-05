namespace SLX_Outlook_AddIn
{
    partial class ContactFound
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
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.txtFirstName = new System.Windows.Forms.TextBox();
            this.txtLastName = new System.Windows.Forms.TextBox();
            this.txtAccount = new System.Windows.Forms.TextBox();
            this.radioButtonCreateOppo = new System.Windows.Forms.RadioButton();
            this.radioButtonCreateTicket = new System.Windows.Forms.RadioButton();
            this.cmdCreate = new System.Windows.Forms.Button();
            this.cmdCancel = new System.Windows.Forms.Button();
            this.radioButtonCreateActivity = new System.Windows.Forms.RadioButton();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(11, 12);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(47, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Contact:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(11, 35);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(50, 13);
            this.label2.TabIndex = 1;
            this.label2.Text = "Account:";
            // 
            // txtFirstName
            // 
            this.txtFirstName.Location = new System.Drawing.Point(64, 5);
            this.txtFirstName.Name = "txtFirstName";
            this.txtFirstName.ReadOnly = true;
            this.txtFirstName.Size = new System.Drawing.Size(100, 20);
            this.txtFirstName.TabIndex = 2;
            // 
            // txtLastName
            // 
            this.txtLastName.Location = new System.Drawing.Point(170, 5);
            this.txtLastName.Name = "txtLastName";
            this.txtLastName.ReadOnly = true;
            this.txtLastName.Size = new System.Drawing.Size(100, 20);
            this.txtLastName.TabIndex = 3;
            // 
            // txtAccount
            // 
            this.txtAccount.Location = new System.Drawing.Point(64, 32);
            this.txtAccount.Name = "txtAccount";
            this.txtAccount.ReadOnly = true;
            this.txtAccount.Size = new System.Drawing.Size(206, 20);
            this.txtAccount.TabIndex = 4;
            // 
            // radioButtonCreateOppo
            // 
            this.radioButtonCreateOppo.AutoSize = true;
            this.radioButtonCreateOppo.Checked = true;
            this.radioButtonCreateOppo.Location = new System.Drawing.Point(14, 69);
            this.radioButtonCreateOppo.Name = "radioButtonCreateOppo";
            this.radioButtonCreateOppo.Size = new System.Drawing.Size(113, 17);
            this.radioButtonCreateOppo.TabIndex = 5;
            this.radioButtonCreateOppo.TabStop = true;
            this.radioButtonCreateOppo.Text = "Create Opportunity";
            this.radioButtonCreateOppo.UseVisualStyleBackColor = true;
            // 
            // radioButtonCreateTicket
            // 
            this.radioButtonCreateTicket.AutoSize = true;
            this.radioButtonCreateTicket.Location = new System.Drawing.Point(14, 92);
            this.radioButtonCreateTicket.Name = "radioButtonCreateTicket";
            this.radioButtonCreateTicket.Size = new System.Drawing.Size(89, 17);
            this.radioButtonCreateTicket.TabIndex = 6;
            this.radioButtonCreateTicket.Text = "Create Ticket";
            this.radioButtonCreateTicket.UseVisualStyleBackColor = true;
            // 
            // cmdCreate
            // 
            this.cmdCreate.Location = new System.Drawing.Point(133, 69);
            this.cmdCreate.Name = "cmdCreate";
            this.cmdCreate.Size = new System.Drawing.Size(75, 23);
            this.cmdCreate.TabIndex = 7;
            this.cmdCreate.Text = "Create";
            this.cmdCreate.UseVisualStyleBackColor = true;
            this.cmdCreate.Click += new System.EventHandler(this.cmdCreate_Click);
            // 
            // cmdCancel
            // 
            this.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.cmdCancel.Location = new System.Drawing.Point(214, 69);
            this.cmdCancel.Name = "cmdCancel";
            this.cmdCancel.Size = new System.Drawing.Size(75, 23);
            this.cmdCancel.TabIndex = 8;
            this.cmdCancel.Text = "Cancel";
            this.cmdCancel.UseVisualStyleBackColor = true;
            // 
            // radioButtonCreateActivity
            // 
            this.radioButtonCreateActivity.AutoSize = true;
            this.radioButtonCreateActivity.Location = new System.Drawing.Point(15, 115);
            this.radioButtonCreateActivity.Name = "radioButtonCreateActivity";
            this.radioButtonCreateActivity.Size = new System.Drawing.Size(93, 17);
            this.radioButtonCreateActivity.TabIndex = 9;
            this.radioButtonCreateActivity.TabStop = true;
            this.radioButtonCreateActivity.Text = "Create Activity";
            this.radioButtonCreateActivity.UseVisualStyleBackColor = true;
            // 
            // ContactFound
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(343, 146);
            this.Controls.Add(this.radioButtonCreateActivity);
            this.Controls.Add(this.cmdCancel);
            this.Controls.Add(this.cmdCreate);
            this.Controls.Add(this.radioButtonCreateTicket);
            this.Controls.Add(this.radioButtonCreateOppo);
            this.Controls.Add(this.txtAccount);
            this.Controls.Add(this.txtLastName);
            this.Controls.Add(this.txtFirstName);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Name = "ContactFound";
            this.Text = "Contact Found";
            this.Load += new System.EventHandler(this.ContactFound_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtFirstName;
        private System.Windows.Forms.TextBox txtLastName;
        private System.Windows.Forms.TextBox txtAccount;
        private System.Windows.Forms.RadioButton radioButtonCreateOppo;
        private System.Windows.Forms.RadioButton radioButtonCreateTicket;
        private System.Windows.Forms.Button cmdCreate;
        private System.Windows.Forms.Button cmdCancel;
        private System.Windows.Forms.RadioButton radioButtonCreateActivity;
    }
}
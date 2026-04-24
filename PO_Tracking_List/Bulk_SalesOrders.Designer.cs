
namespace PO_Tracking_List
{
    partial class Bulk_SalesOrders
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
            this.panel1 = new System.Windows.Forms.Panel();
            this.btn_Select_All = new System.Windows.Forms.Button();
            this.btn_SO_Confirm = new System.Windows.Forms.Button();
            this.textselect = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.lbl_bulk_item = new System.Windows.Forms.Label();
            this.chk = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Controls.Add(this.btn_Select_All);
            this.panel1.Controls.Add(this.btn_SO_Confirm);
            this.panel1.Controls.Add(this.textselect);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Location = new System.Drawing.Point(1, 2);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(397, 75);
            this.panel1.TabIndex = 0;
            // 
            // btn_Select_All
            // 
            this.btn_Select_All.Font = new System.Drawing.Font("Microsoft Tai Le", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_Select_All.Location = new System.Drawing.Point(24, 42);
            this.btn_Select_All.Name = "btn_Select_All";
            this.btn_Select_All.Size = new System.Drawing.Size(75, 28);
            this.btn_Select_All.TabIndex = 3;
            this.btn_Select_All.Text = "SelectAll";
            this.btn_Select_All.UseVisualStyleBackColor = true;
            this.btn_Select_All.Click += new System.EventHandler(this.btn_Select_All_Click);
            // 
            // btn_SO_Confirm
            // 
            this.btn_SO_Confirm.Font = new System.Drawing.Font("Microsoft Tai Le", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_SO_Confirm.Location = new System.Drawing.Point(257, 12);
            this.btn_SO_Confirm.Name = "btn_SO_Confirm";
            this.btn_SO_Confirm.Size = new System.Drawing.Size(66, 23);
            this.btn_SO_Confirm.TabIndex = 2;
            this.btn_SO_Confirm.Text = "Confirm";
            this.btn_SO_Confirm.UseVisualStyleBackColor = true;
            this.btn_SO_Confirm.Click += new System.EventHandler(this.btn_SO_Confirm_Click);
            // 
            // textselect
            // 
            this.textselect.Location = new System.Drawing.Point(102, 12);
            this.textselect.Multiline = true;
            this.textselect.Name = "textselect";
            this.textselect.Size = new System.Drawing.Size(139, 23);
            this.textselect.TabIndex = 1;
            this.textselect.TextChanged += new System.EventHandler(this.textselect_TextChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Tai Le", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(24, 15);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(72, 16);
            this.label1.TabIndex = 0;
            this.label1.Text = "Sales Order";
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.BackgroundColor = System.Drawing.SystemColors.ButtonHighlight;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.chk});
            this.dataGridView1.Location = new System.Drawing.Point(1, 80);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dataGridView1.Size = new System.Drawing.Size(397, 476);
            this.dataGridView1.TabIndex = 1;
            // 
            // lbl_bulk_item
            // 
            this.lbl_bulk_item.AutoSize = true;
            this.lbl_bulk_item.Font = new System.Drawing.Font("Microsoft Tai Le", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_bulk_item.Location = new System.Drawing.Point(29, 570);
            this.lbl_bulk_item.Name = "lbl_bulk_item";
            this.lbl_bulk_item.Size = new System.Drawing.Size(16, 14);
            this.lbl_bulk_item.TabIndex = 2;
            this.lbl_bulk_item.Text = "...";
            // 
            // chk
            // 
            this.chk.FalseValue = "0";
            this.chk.HeaderText = "Select";
            this.chk.MinimumWidth = 4;
            this.chk.Name = "chk";
            this.chk.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.chk.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            this.chk.TrueValue = "1";
            this.chk.Width = 50;
            // 
            // Bulk_SalesOrders
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(399, 595);
            this.Controls.Add(this.lbl_bulk_item);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.panel1);
            this.Name = "Bulk_SalesOrders";
            this.Text = "Bulk_SalesOrders";
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Button btn_Select_All;
        private System.Windows.Forms.Button btn_SO_Confirm;
        private System.Windows.Forms.TextBox textselect;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label lbl_bulk_item;
        private System.Windows.Forms.DataGridViewCheckBoxColumn chk;
    }
}
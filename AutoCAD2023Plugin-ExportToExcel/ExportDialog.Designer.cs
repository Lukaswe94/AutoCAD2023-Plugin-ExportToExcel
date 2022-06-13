
namespace AutoCAD2023Plugin.ExportToExcel
{
    partial class ExportDialog
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;
        private System.Windows.Forms.Button export;
        private System.Windows.Forms.Button abort;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.DataGridView gridView;

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
            this.export = new System.Windows.Forms.Button();
            this.abort = new System.Windows.Forms.Button();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.tableLayoutPanel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // export
            // 
            this.export.Location = new System.Drawing.Point(103, 301);
            this.export.Name = "export";
            this.export.Size = new System.Drawing.Size(75, 23);
            this.export.TabIndex = 1;
            this.export.Text = "Exportieren";
            this.export.UseVisualStyleBackColor = true;
            this.export.Click += new System.EventHandler(ExportButton_Click);
            // 
            // abort
            // 
            this.abort.Location = new System.Drawing.Point(218, 301);
            this.abort.Name = "abort";
            this.abort.Size = new System.Drawing.Size(75, 23);
            this.abort.TabIndex = 2;
            this.abort.Text = "Abbrechen";
            this.abort.UseVisualStyleBackColor = true;
            this.abort.Click += new System.EventHandler(AbortButton_Click);
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.AutoSize = true;
            this.tableLayoutPanel1.ColumnCount = 3;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 100F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 40.4908F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 59.5092F));
            this.tableLayoutPanel1.Controls.Add(this.abort, 2, 1);
            this.tableLayoutPanel1.Controls.Add(this.export, 1, 1);
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 2;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10.90909F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(386, 331);
            this.tableLayoutPanel1.TabIndex = 3;
            // 
            // ExportDialog
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(387, 331);
            this.Controls.Add(this.tableLayoutPanel1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ExportDialog";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "ExportDialog";
            this.tableLayoutPanel1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

    }
}
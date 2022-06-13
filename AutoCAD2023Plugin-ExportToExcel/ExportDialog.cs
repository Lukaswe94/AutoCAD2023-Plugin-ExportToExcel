using System;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace AutoCAD2023Plugin.ExportToExcel
{
    public partial class ExportDialog : Form
    {
        public ExportDialog()
        {
            InitializeComponent();
        }
        public void AddControl(int row, int col, int columnSpan, Control control)
        {
            if (control is DataGridView view) gridView = view;
            tableLayoutPanel1.Controls.Add(control);
            tableLayoutPanel1.SetColumn(control, col);
            tableLayoutPanel1.SetRow(control, row);
            tableLayoutPanel1.SetColumnSpan(control, columnSpan);
        }
        public void ExportButton_Click(object sender, EventArgs e)
        {
            gridView.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            gridView.MultiSelect = true;
            gridView.SelectAll();
            DataObject obj = gridView.GetClipboardContent();
            Clipboard.SetDataObject(obj);
            Application excel = new Application
            {
                Visible = false,
                DisplayAlerts = false
            };

            Workbook workbook = excel.Workbooks.Add(Type.Missing);
            Worksheet worksheet = null;
            worksheet = (Worksheet)workbook.ActiveSheet;
            Range range = (Range)worksheet.Cells[1, 1];
            range.Select();
            worksheet.PasteSpecial(range, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
            worksheet.SaveAs(String.Format("Attribute{0}{1}.xlsx", gridView.Name, DateTime.Now.ToString("MM-dd-yyyy HH.mm.ss")), 
                Type.Missing, Type.Missing, Type.Missing, false, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
            Clipboard.Clear();
            gridView.ClearSelection();
            workbook.Close();
            excel.Quit();
            this.DialogResult = DialogResult.OK;
            this.Close();

        }
        public void AbortButton_Click(object sender, EventArgs e) 
        {
            DialogResult = DialogResult.Abort;
            this.Close();
        }

    }
}

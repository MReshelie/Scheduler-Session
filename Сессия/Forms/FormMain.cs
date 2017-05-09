using DevExpress.Spreadsheet;
using DevExpress.XtraBars.Docking2010;
using DevExpress.XtraEditors;
using System;
using System.IO;
using System.Linq;


namespace Сессия
{
    public partial class FormMain : DevExpress.XtraEditors.XtraForm
    {
        public FormMain()
        {
            InitializeComponent();
        }

        private void FormMain_Load(object sender, EventArgs e)
        {
            if (!UIButtons.UIButtonsEnabled(windowsUIButtonPanelExcel, false)) Console.WriteLine("Что то пошло не так!!!");
        }

        private void tileBar_SelectedItemChanged(object sender, TileItemEventArgs e)
        {
            navigationFrame.SelectedPageIndex = tileBarGroupTables.Items.IndexOf(e.Item);
        }

        private void windowsUIButtonPanelExcel_ButtonClick(object sender, ButtonEventArgs e)
        {
            string btCaption = ((WindowsUIButton)e.Button).Caption.ToString();

            switch (btCaption)
            {
                case "Открыть":
                    if (openFileDialogExcel.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                        using (FileStream stream = new FileStream(openFileDialogExcel.FileName, FileMode.Open))
                        {
                            spreadsheetControlExcel.LoadDocument(stream, DevExpress.Spreadsheet.DocumentFormat.OpenXml);

                            if (!UIButtons.UIButtonsEnabled(windowsUIButtonPanelExcel, true)) Console.WriteLine("Что то пошло не так!!!");
                        }

                    break;
                case "Обработать":
                    break;
                case "Очистить":
                    IWorkbook workbook = spreadsheetControlExcel.Document;
                    WorksheetCollection worksheets = workbook.Worksheets;

                    int i = worksheets.Count - 1;
                    if (i > 0)
                        do
                        {
                            workbook.Worksheets.RemoveAt(i);
                            i--;
                        } while (i > 0);

                    workbook.Worksheets.Insert(0, "Лист1");
                    workbook.Worksheets.ActiveWorksheet = workbook.Worksheets[1];
                    workbook.Worksheets.RemoveAt(1);

                    if (!UIButtons.UIButtonsEnabled(windowsUIButtonPanelExcel, false)) Console.WriteLine("Что то пошло не так!!!"); break;
            }
        }
    }
}
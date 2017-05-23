using DevExpress.Spreadsheet;
using DevExpress.Spreadsheet.Export;
using DevExpress.Utils;
using DevExpress.XtraBars.Docking2010;
using DevExpress.XtraEditors;
using DevExpress.XtraGrid.Views.Base;
using DevExpress.XtraScheduler;
using DevExpress.XtraScheduler.Drawing;
using DevExpress.XtraSplashScreen;
using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Сессия.Classes;
using Сессия.Database;

namespace Сессия
{
    public partial class FormMain : DevExpress.XtraEditors.XtraForm
    {
        private int LastSplitContainerControlSplitterPosition;
        string[] connectionString = Utilities.GetConnectionStrings();
        int iCS = 0, id = 0;
        Image resImage;

        public FormMain()
        {
            InitializeComponent();
            this.schedulerControlTimeTable.ToolTipController = toolTipControllerScheduler;

            BaseInitialization();
            LoadData();
        }

        void BaseInitialization()
        {
            SplashImagePainter.Painter.ViewInfo.Stage = "Загрузка компонентов";
            for (int i = 1; i <= 100; i++)
            {
                System.Threading.Thread.Sleep(20);
                SplashImagePainter.Painter.ViewInfo.Counter = i;
                SplashScreenManager.Default.Invalidate();
            }

            #region #toolTipControllerScheduler
            resImage = Image.FromFile(@"..\..\Resources\appointment.gif");
            this.toolTipControllerScheduler.ShowBeak = true;
            this.schedulerControlTimeTable.OptionsView.ToolTipVisibility = ToolTipVisibility.Always;
            this.toolTipControllerScheduler.ToolTipType = ToolTipType.SuperTip;
            #endregion

            #region #AppointmentEvents
            schedulerStorage1.AppointmentsInserted += new PersistentObjectsEventHandler(schedulerStorage1_AppointmentsInserted);
            schedulerStorage1.AppointmentsChanged += new PersistentObjectsEventHandler(schedulerStorage1_AppointmentsChanged);
            schedulerStorage1.AppointmentsDeleted += new PersistentObjectsEventHandler(schedulerStorage1_AppointmentsDeleted);
            #endregion #AppointmentEvents

            #region #AppointmentDependencyEvents
            schedulerStorage1.AppointmentDependenciesInserted += new PersistentObjectsEventHandler(schedulerStorage1_AppointmentDependenciesInserted);
            schedulerStorage1.AppointmentDependenciesChanged += new PersistentObjectsEventHandler(schedulerStorage1_AppointmentDependenciesChanged);
            schedulerStorage1.AppointmentDependenciesDeleted += new PersistentObjectsEventHandler(schedulerStorage1_AppointmentDependenciesDeleted);
            #endregion #AppointmentDependencyEvents

            #region #schedulerControlTimeTable Events
            //Fix the view type and splitter position.
            this.schedulerControlTimeTable.ActiveViewChanged += new EventHandler(schedulerControlTimeTable_ActiveViewChanged);
            // Set the date to show existing appointments from the database.
            this.schedulerControlTimeTable.InitNewAppointment += new AppointmentEventHandler(schedulerControlTimeTable_InitNewAppointment);
            this.schedulerControlTimeTable.AppointmentViewInfoCustomizing += schedulerControlTimeTable_AppointmentViewInfoCustomizing;
            #endregion

            #region #splitContainerControlScheduler
            this.splitContainerControlScheduler.SplitterPositionChanged += new System.EventHandler(this.splitContainerControlScheduler_SplitterPositionChanged);
            #endregion

            #region  #scales TimelineView & GanttView
            TimeScaleCollection scalesTlV = schedulerControlTimeTable.TimelineView.Scales;
            scalesTlV.BeginUpdate();
            try
            {
                scalesTlV.Clear();
                scalesTlV.Add(new MyTimeScaleDay(TimeSpan.FromDays(1)));
                scalesTlV.Add(new MyTimeScaleLessThanDay(TimeSpan.FromHours(1)));
                scalesTlV.Add(new MyTimeScaleLessThanDay(TimeSpan.FromMinutes(30)));
                scalesTlV.Add(new MyTimeScaleLessThanDay(TimeSpan.FromMinutes(15)));
            }
            finally
            {
                scalesTlV.EndUpdate();
            }

            TimeScaleCollection scalesGV = schedulerControlTimeTable.GanttView.Scales;
            scalesGV.BeginUpdate();
            try
            {
                scalesGV.Clear();
                scalesGV.Add(new MyTimeScaleDay(TimeSpan.FromDays(1)));
                scalesGV.Add(new MyTimeScaleLessThanDay(TimeSpan.FromHours(1)));
                scalesGV.Add(new MyTimeScaleLessThanDay(TimeSpan.FromMinutes(30)));
                scalesGV.Add(new MyTimeScaleLessThanDay(TimeSpan.FromMinutes(15)));
            }
            finally
            {
                scalesGV.EndUpdate();
            }
            #endregion #scales

            #region #Adjustment     
            schedulerControlTimeTable.Start = DateTime.Today;
            schedulerControlTimeTable.ActiveViewType = DevExpress.XtraScheduler.SchedulerViewType.Timeline;
            schedulerControlTimeTable.GroupType = SchedulerGroupType.Resource;
            schedulerControlTimeTable.GanttView.CellsAutoHeightOptions.Enabled = true;
            // Hide unnecessary visual elements.
            //schedulerControlTimeTable.GanttView.ShowResourceHeaders = false;
            //schedulerControlTimeTable.GanttView.NavigationButtonVisibility = NavigationButtonVisibility.Never;
            // Disable user sorting in the Resource Tree (clicking the column will not change the sort order).
            колDescription.OptionsColumn.AllowSort = false;
            #endregion #Adjustment
        }

        void LoadData()
        {
            SplashImagePainter.Painter.ViewInfo.Counter = 0;
            SplashImagePainter.Painter.ViewInfo.Stage = "Подключение к БД";
            SplashScreenManager.Default.Invalidate();

            GetDataSet();

            #region #ConnectionString
            for (int i = 0; i < connectionString.Length; i++)
            {
                string subString = "Сессия.Properties.Settings.SessionDBlConnectionString";

                if (connectionString[i].Trim() == subString)
                {
                    connectionString[i] = ConfigurationManager.ConnectionStrings[connectionString[i].Trim()].ConnectionString;
                    iCS = i;
                }
            }
            #endregion

            #region #CommitIdToDataSource
            schedulerStorage1.Appointments.CommitIdToDataSource = false;
            this.appointmentsTableAdapter.Adapter.RowUpdated += new SqlRowUpdatedEventHandler(appointmentsTableAdapter_RowUpdated);
            #endregion #CommitIdToDataSource
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            SplashScreenManager.HideImage();
        }

        private void FormMain_Load(object sender, EventArgs e)
        {
            SessionDBlDataSet.WorksheetDataDataTable tabWsDataTable = worksheetDataTableAdapter.GetData();
            var q = from c in tabWsDataTable.AsEnumerable() select c.IdRow;

            if (q.Count() > 0)
            {
                if (!UIButtons.UIButtonsEnabled(windowsUIButtonPanelExcel, new bool[] { true, true, false, false, false, false, false, false, false, false })) Console.WriteLine("Что то пошло не так!!!");
            }
            else
                if (!UIButtons.UIButtonsEnabled(windowsUIButtonPanelExcel, new bool[] { true, false, false, false, false, false, false, false, false, false })) Console.WriteLine("Что то пошло не так!!!");
        }

        #region Работа с БД
        private void appointmentsTableAdapter_RowUpdated(object sender, SqlRowUpdatedEventArgs e)
        {
            if (e.Status == UpdateStatus.Continue && e.StatementType == StatementType.Insert)
            {
                id = 0;
                using (SqlCommand cmd = new SqlCommand("SELECT @@IDENTITY", appointmentsTableAdapter.Connection))
                {
                    id = Convert.ToInt32(cmd.ExecuteScalar());
                    e.Row["UniqueId"] = id;
                }
            }
        }
        #endregion

        #region Основное меню tileBar
        private void tileBar_SelectedItemChanged(object sender, TileItemEventArgs e)
        {
            navigationFrame.SelectedPageIndex = tileBarGroupTables.Items.IndexOf(e.Item);
        }
        #endregion

        #region Боковое меню Импорта из Excel
        private void windowsUIButtonPanelExcel_ButtonClick(object sender, ButtonEventArgs e)
        {
            if (e.Button.GetType().Name == "WindowsUISeparator") return;

            IWorkbook workbook = spreadsheetControlExcel.Document;
            WorksheetCollection worksheets = workbook.Worksheets;
            Range usedRange;
            int i, iRow, jRow;
            DateTime dDay;
            string sEndTime, btCaption = ((WindowsUIButton)e.Button).Caption.ToString();
            string sSql = string.Empty;

            switch (btCaption)
            {
                case "Открыть файл":
                    if (openFileDialogExcel.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        SplashScreenManager.ShowForm(typeof(Сессия.Forms.WaitFormUpdate));
                        SplashScreenManager.Default.SetWaitFormCaption("Ожидайте пожалуйста.");
                        SplashScreenManager.Default.SetWaitFormDescription("Загрузка исходного файла");

                        using (FileStream stream = new FileStream(openFileDialogExcel.FileName, FileMode.Open))
                        {
                            spreadsheetControlExcel.LoadDocument(stream, DevExpress.Spreadsheet.DocumentFormat.OpenXml);

                            if (!UIButtons.UIButtonsEnabled(windowsUIButtonPanelExcel, new bool[] { false, false, false, true, true, false, false, false, false, true })) Console.WriteLine("Что то пошло не так!!!");
                        }

                        workbook.Worksheets[0].Name = "МАДИ";
                        workbook.Worksheets[0].Columns.AutoFit(1, 5);
                        SplashScreenManager.CloseForm(false);
                    }

                    break;
                case "Загрузить из БД":
                    SplashScreenManager.ShowForm(typeof(Сессия.Forms.WaitFormUpdate));
                    SplashScreenManager.Default.SetWaitFormCaption("Ожидайте пожалуйста.");
                    SplashScreenManager.Default.SetWaitFormDescription("Загрузка листов из БД");

                    byte[] receivedBytes;
                    using (SqlConnection connection = new SqlConnection(connectionString[iCS]))
                    {
                        connection.Open();
                        SqlCommand command = connection.CreateCommand();
                        command.CommandText = "SELECT Data FROM WorksheetData WHERE IdRow = (SELECT MAX(IdRow) FROM WorksheetData)";
                        SqlDataReader sqlReader = command.ExecuteReader();

                        if (sqlReader.Read())
                            receivedBytes = (byte[])sqlReader[0];
                        else
                        {
                            SplashScreenManager.CloseForm(false);

                            if (XtraMessageBox.Show("В БД нет ранее сохраненных данных.", "Загрузка данных", MessageBoxButtons.OK, MessageBoxIcon.Warning) == DialogResult.OK)
                                if (!UIButtons.UIButtonsEnabled(windowsUIButtonPanelExcel, new bool[] { true, false, false, false, false, false, false, false, false, false })) Console.WriteLine("Что то пошло не так!!!");

                            break;
                        }
                    }

                    workbook.LoadDocument(receivedBytes, DevExpress.Spreadsheet.DocumentFormat.OpenXml);

                    switch (worksheets.Count)
                    {
                        case 1:
                            if (!UIButtons.UIButtonsEnabled(windowsUIButtonPanelExcel, new bool[] { false, false, false, true, true, false, false, false, false, true })) Console.WriteLine("Что то пошло не так!!!");

                            break;
                        case 2:
                            if (sessionDBlDataSet.Appointments.Count() > 0)
                            {
                                if (!UIButtons.UIButtonsEnabled(windowsUIButtonPanelExcel, new bool[] { false, false, false, false, false, false, false, true, false, true })) Console.WriteLine("Что то пошло не так!!!");
                            }
                            else
                                if (!UIButtons.UIButtonsEnabled(windowsUIButtonPanelExcel, new bool[] { false, false, false, false, false, false, true, false, false, true })) Console.WriteLine("Что то пошло не так!!!");

                            break;
                        case 3:
                            if (!UIButtons.UIButtonsEnabled(windowsUIButtonPanelExcel, new bool[] { false, false, false, false, false, false, false, true, false, true })) Console.WriteLine("Что то пошло не так!!!");

                            break;
                    }

                    SplashScreenManager.CloseForm(false);

                    break;
                case "Обработать":
                    SplashScreenManager.ShowForm(this, typeof(Сессия.Forms.WaitFormUpdate), true, true, false);
                    SplashScreenManager.Default.SetWaitFormCaption("Ожидайте пожалуйста.");

                    workbook.Worksheets.Insert(1, "Расписание");
                    workbook.Worksheets[1].Cells["A1"].Value = "Type";
                    workbook.Worksheets[1].Cells["B1"].Value = "StartDate";
                    workbook.Worksheets[1].Cells["C1"].Value = "EndDate";
                    workbook.Worksheets[1].Cells["D1"].Value = "AllDay";
                    workbook.Worksheets[1].Cells["E1"].Value = "Subject";
                    workbook.Worksheets[1].Cells["F1"].Value = "Location";
                    workbook.Worksheets[1].Cells["G1"].Value = "Description";
                    workbook.Worksheets[1].Cells["H1"].Value = "Status";
                    workbook.Worksheets[1].Cells["I1"].Value = "Label";
                    workbook.Worksheets[1].Cells["J1"].Value = "ResourceId";
                    workbook.Worksheets[1].Cells["K1"].Value = "PercentComplete";
                    workbook.Worksheets[1].Cells["L1"].Value = "TimeZoneId";

                    usedRange = workbook.Worksheets[0].GetUsedRange();
                    iRow = 0;
                    jRow = 1;

                    do
                    {
                        SplashScreenManager.Default.SetWaitFormDescription(string.Format("Обработано {0:N0}% данных", (((double)iRow / (double)usedRange.RowCount) * 100)));

                        dDay = Convert.ToDateTime(workbook.Worksheets[0].GetCellValue(0, iRow).ToString().Trim()).Date;
                        iRow += 2;

                        while (workbook.Worksheets[0].GetCellValue(0, iRow).ToString().Trim().Length > 0)
                        {
                            workbook.Worksheets[1].Cells[jRow, 0].Value = 0;
                            workbook.Worksheets[1].Cells[jRow, 1].Value = $"{dDay:d} {(workbook.Worksheets[0].GetCellValue(1, iRow).ToString().Trim())}";
                            sEndTime = (int.Parse(workbook.Worksheets[0].GetCellValue(1, iRow).ToString().Trim().Substring(0, 2)) + 4).ToString() + ":" +
                                                                           workbook.Worksheets[0].GetCellValue(1, iRow).ToString().Trim().Substring(3, 2);
                            workbook.Worksheets[1].Cells[jRow, 2].Value = $"{dDay:d} {sEndTime}";
                            workbook.Worksheets[1].Cells[jRow, 3].Value = 0;
                            workbook.Worksheets[1].Cells[jRow, 4].Value = workbook.Worksheets[0].GetCellValue(0, iRow).ToString().Trim() + "; Экзамен";
                            workbook.Worksheets[1].Cells[jRow, 5].Value = workbook.Worksheets[0].GetCellValue(4, iRow).ToString().Trim();
                            workbook.Worksheets[1].Cells[jRow, 6].Value = workbook.Worksheets[0].GetCellValue(2, iRow).ToString().Trim();
                            workbook.Worksheets[1].Cells[jRow, 7].Value = 2;
                            workbook.Worksheets[1].Cells[jRow, 8].Value = 3;

                            using (SqlConnection connection = new SqlConnection(connectionString[iCS]))
                            {
                                connection.Open();
                                SqlCommand command = connection.CreateCommand();
                                command.CommandText = string.Empty;
                                command.CommandText = "SELECT [Id] FROM Resources WHERE ([Description] Like N'" + workbook.Worksheets[0].GetCellValue(3, iRow).ToString().Trim() + "');";
                                workbook.Worksheets[1].Cells[jRow, 9].Value = Int32.Parse(Convert.ToString(command.ExecuteScalar()));
                            }

                            // Придумать как читать процент до завершения при проверке времени проведения экзамена если день совпадает
                            if (Convert.ToDateTime(workbook.Worksheets[1].GetCellValue(2, jRow).ToString().Trim()) < DateTime.Today)
                                workbook.Worksheets[1].Cells[jRow, 10].Value = 100;
                            else
                                workbook.Worksheets[1].Cells[jRow, 10].Value = 0;

                            workbook.Worksheets[1].Cells[jRow, 11].Value = "Russian Standard Time";
                            workbook.Worksheets[1].ScrollToRow(jRow);
                            iRow += 1;
                            jRow += 1;
                        }

                        iRow += 1;
                    } while (iRow <= usedRange.RowCount);

                    workbook.Worksheets[1].ScrollToRow(0);
                    workbook.Worksheets[1].Columns.AutoFit(0, 11);

                    if (!UIButtons.UIButtonsEnabled(windowsUIButtonPanelExcel, new bool[] { false, false, false, false, true, false, true, false, false, true })) Console.WriteLine("Что то пошло не так!!!");

                    SplashScreenManager.CloseForm(false);

                    break;
                case "Сохранить в БД":
                    SplashScreenManager.ShowForm(typeof(Сессия.Forms.WaitFormUpdate));
                    SplashScreenManager.Default.SetWaitFormCaption("Ожидайте пожалуйста.");
                    SplashScreenManager.Default.SetWaitFormDescription("Запись листов в БД");

                    using (SqlConnection connection = new SqlConnection(connectionString[iCS]))
                    {
                        connection.Open();
                        SqlCommand command = connection.CreateCommand();
                        command.CommandText = "TRUNCATE TABLE WorksheetData;";
                        command.ExecuteNonQuery();
                        command.CommandText = "INSERT INTO WorksheetData(Data) VALUES(@Data)";
                        SqlParameter dataParameter = new SqlParameter("@Data", SqlDbType.VarBinary);
                        dataParameter.Value = spreadsheetControlExcel.SaveDocument(DevExpress.Spreadsheet.DocumentFormat.OpenXml);
                        command.Parameters.Add(dataParameter);
                        command.ExecuteNonQuery();
                    }

                    switch (worksheets.Count)
                    {
                        case 1:
                            if (!UIButtons.UIButtonsEnabled(windowsUIButtonPanelExcel, new bool[] { false, false, false, true, true, false, false, false, false, true })) Console.WriteLine("Что то пошло не так!!!");

                            break;
                        case 2:
                            if (!UIButtons.UIButtonsEnabled(windowsUIButtonPanelExcel, new bool[] { false, false, false, false, false, false, true, false, false, true })) Console.WriteLine("Что то пошло не так!!!");

                            break;
                        case 3:
                            if (!UIButtons.UIButtonsEnabled(windowsUIButtonPanelExcel, new bool[] { false, false, false, false, false, false, false, true, false, true })) Console.WriteLine("Что то пошло не так!!!");

                            break;
                    }

                    SplashScreenManager.CloseForm(false);

                    break;
                case "Расписание":
                    SplashScreenManager.ShowForm(typeof(Сессия.Forms.WaitFormUpdate));
                    SplashScreenManager.Default.SetWaitFormCaption("Ожидайте пожалуйста.");
                    SplashScreenManager.Default.SetWaitFormDescription("Запись данных в таблицу БД");

                    Worksheet worksheet = spreadsheetControlExcel.Document.Worksheets[1];
                    usedRange = workbook.Worksheets[1].GetUsedRange();

                    // Create a data table with column names obtained from the first row in a range.
                    // Column data types are obtained from cell value types of cells in the first data row of the worksheet range.
                    DataTable dataTable = worksheet.CreateDataTable(usedRange, true);

                    // Change the data type of the "As Of" column to text.
                    // Create the exporter that obtains data from the specified range and populates the specified data table. 
                    DataTableExporter exporter = worksheet.CreateDataTableExporter(usedRange, dataTable, true);

                    // Handle value conversion errors.
                    exporter.CellValueConversionError += exporter_CellValueConversionError;

                    // Specify a custom converter for the "As Of" column.
                    DateTimeToStringConverter toDateStringConverter = new DateTimeToStringConverter();

                    // Set the export value for empty cell.
                    // Specify that empty cells and cells with errors should be processed.
                    exporter.Options.ConvertEmptyCells = true;
                    exporter.Options.DefaultCellValueToColumnTypeConverter.SkipErrorValues = false;

                    // Perform the export.
                    exporter.Export();

                    // A custom method that displays the resulting data table.
                    //ShowResult(dataTable);

                    // Add records to the database
                    using (SqlConnection connection = new SqlConnection(connectionString[iCS]))
                    {
                        connection.Open();
                        SqlCommand command = connection.CreateCommand();
                        command.CommandText = "TRUNCATE TABLE Appointments;";
                        command.ExecuteNonQuery();
                    }

                    using (SqlConnection connection = new SqlConnection(connectionString[iCS]))
                    {
                        connection.Open();
                        SqlBulkCopy objBulk = new SqlBulkCopy(connection);
                        objBulk.DestinationTableName = "Appointments";
                        objBulk.ColumnMappings.Add("Type", "Type");
                        objBulk.ColumnMappings.Add("StartDate", "StartDate");
                        objBulk.ColumnMappings.Add("EndDate", "EndDate");
                        objBulk.ColumnMappings.Add("AllDay", "AllDay");
                        objBulk.ColumnMappings.Add("Subject", "Subject");
                        objBulk.ColumnMappings.Add("Location", "Location");
                        objBulk.ColumnMappings.Add("Description", "Description");
                        objBulk.ColumnMappings.Add("Status", "Status");
                        objBulk.ColumnMappings.Add("Label", "Label");
                        objBulk.ColumnMappings.Add("ResourceId", "ResourceId");
                        objBulk.ColumnMappings.Add("PercentComplete", "PercentComplete");
                        objBulk.ColumnMappings.Add("TimeZoneId", "TimeZoneId");
                        objBulk.WriteToServer(dataTable);
                    }

                    // Send the update data to the database
                    try
                    {
                        appointmentsTableAdapter.Update(sessionDBlDataSet.Appointments);
                    }
                    catch (Exception ex)
                    {
                        string message = string.Format("Нет возможности обновить данные в таблице БД.\n{0}", ex.Message);
                        MessageBox.Show(message, "Ошибка записи", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                    GetDataSet();

                    if (!UIButtons.UIButtonsEnabled(windowsUIButtonPanelExcel, new bool[] { false, false, false, false, true, false, false, true, false, true })) Console.WriteLine("Что то пошло не так!!!");

                    SplashScreenManager.CloseForm(false);

                    break;
                case "Консультации":
                    SplashScreenManager.ShowForm(typeof(Сессия.Forms.WaitFormUpdate));
                    SplashScreenManager.Default.SetWaitFormCaption("Ожидайте пожалуйста.");

                    if (worksheets.Count == 3)
                        workbook.Worksheets.RemoveAt(worksheets.Count - 1);

                    AppointmentLabelCollection arrLabel = this.schedulerStorage1.Appointments.Labels;
                    string[] words;
                    string sRange = string.Empty;
                    DateTime dDayCur = DateTime.Today;

                    workbook.Worksheets.Insert(2, "Консультации");
                    workbook.Worksheets[2].Cells["A1"].Value = "Группа";
                    workbook.Worksheets[2].Cells["B1"].Value = "Время";
                    workbook.Worksheets[2].Cells["C1"].Value = "Аттестация";
                    workbook.Worksheets[2].Cells["D1"].Value = "Дисциплина";
                    workbook.Worksheets[2].Cells["E1"].Value = "Аудит.";
                    workbook.Worksheets[2].Cells["F1"].Value = "Преподаватель";
                    workbook.Worksheets[2].Cells["G1"].Value = "Дата";
                    workbook.Worksheets[2].Cells["H1"].Value = "Время";
                    workbook.Worksheets[2].Cells["I1"].Value = "Аудит.";
                    workbook.Worksheets[2].Cells["J1"].Value = "Вид.";

                    var qExsam = sessionDBlDataSet.Appointments.OrderBy(x => x.StartDate.Date).ThenBy(x => x.ResourceId);

                    if (qExsam.Count() > 0)
                    {
                        jRow = 1;
                        foreach (var rExsam in qExsam)
                        {
                            if (jRow == 1)
                            {
                                dDayCur = rExsam.StartDate.Date;
                                workbook.Worksheets[2].Cells[jRow, 0].Value = dDayCur; //.ToString("dd MMMM yyy - dddd");
                                jRow++;
                            }
                            else
                                if (dDayCur.Date != rExsam.StartDate.Date)
                            {
                                dDayCur = rExsam.StartDate;
                                workbook.Worksheets[2].Cells[jRow, 0].Value = dDayCur;//.ToString("dd MMMM yyy - dddd");
                                jRow++;
                            }

                            words = rExsam.Subject.Split(';');

                            workbook.Worksheets[2].Cells[jRow, 0].Value = words[0].Trim();
                            workbook.Worksheets[2].Cells[jRow, 1].Value = rExsam.StartDate.ToString("HH:mm");

                            workbook.Worksheets[2].Cells[jRow, 2].Value = arrLabel.GetByIndex(rExsam.Label).ToString();

                            workbook.Worksheets[2].Cells[jRow, 3].Value = rExsam.Description;

                            workbook.Worksheets[2].Cells[jRow, 4].Value = sessionDBlDataSet.Resources.Where(n => n.Id == rExsam.ResourceId).Select(n => n.Description).First();
                            workbook.Worksheets[2].Cells[jRow, 5].Value = rExsam.Location;

                            // Поиск консультации
                            var qCons = qExsam.Where(m => m.StartDate.Date < rExsam.StartDate.Date).
                                               Where(m => m.Location == rExsam.Location).
                                               Where(m => m.Subject.Contains(words[0].Trim()) == true).
                                               Select(m => new { m.StartDate, m.ResourceId, m.Location, m.Label });

                            if (qCons.Count() > 0)
                                foreach (var rCons in qCons)
                                    if (rCons.Label > 3)
                                    {
                                        workbook.Worksheets[2].Cells[jRow, 6].Value = rCons.StartDate.Date;
                                        workbook.Worksheets[2].Cells[jRow, 7].Value = rCons.StartDate.ToString("HH:mm");
                                        workbook.Worksheets[2].Cells[jRow, 8].Value = sessionDBlDataSet.Resources.Where(n => n.Id == rCons.ResourceId).Select(n => n.Description).First();
                                        workbook.Worksheets[2].Cells[jRow, 9].Value = arrLabel.GetByIndex(rCons.Label).ToString();
                                    }

                            SplashScreenManager.Default.SetWaitFormDescription(string.Format("Обработана {0:N0}-я строка", jRow));
                            jRow++;
                        }
                    }

                    workbook.Worksheets[2].ScrollToRow(1);
                    workbook.Worksheets[2].Range["A1:J1"].Font.Size = 14;
                    workbook.Worksheets[2].Range["A1:J1"].Font.FontStyle = SpreadsheetFontStyle.Bold;
                    workbook.Worksheets[2].Columns.AutoFit(0, 9);

                    usedRange = workbook.Worksheets[2].GetUsedRange();
                    for (jRow = 0; jRow < usedRange.RowCount; jRow++)
                        if (workbook.Worksheets[2].GetCellValue(0, jRow).IsDateTime)
                        {
                            sRange = "A" + (jRow + 1).ToString().Trim() + ":J" + (jRow + 1).ToString().Trim();
                            dDayCur = workbook.Worksheets[2].GetCellValue(0, jRow).DateTimeValue;
                            workbook.Worksheets[2].MergeCells(workbook.Worksheets[2].Range[sRange]);
                            workbook.Worksheets[2].Cells[jRow, 0].Value = dDayCur.ToString("dd MMMM yyy - dddd");
                            workbook.Worksheets[2].Range[sRange].Font.FontStyle = SpreadsheetFontStyle.Bold;
                            workbook.Worksheets[2].Range[sRange].Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center;
                            SplashScreenManager.Default.SetWaitFormDescription(string.Format("Обработана {0:N0}-я строка", jRow));
                        }

                    usedRange.Borders.SetAllBorders(Color.Black, BorderLineStyle.Thick);
                    usedRange.SetInsideBorders(Color.Black, BorderLineStyle.Thin);
                    workbook.Worksheets[2].Range["A1:J1"].Borders.SetOutsideBorders(Color.Black, BorderLineStyle.Thick);

                    workbook.Worksheets[2].ScrollToRow(1);
                    workbook.Worksheets[2].Rows["1"].Insert();
                    workbook.Worksheets[2].MergeCells(workbook.Worksheets[2].Range["A1:J1"]);
                    workbook.Worksheets[2].Cells["A1"].Value = "Московский автомобильно-дорожный государственный технический университет (МАДИ)";
                    workbook.Worksheets[2].Rows["2"].Insert();
                    workbook.Worksheets[2].MergeCells(workbook.Worksheets[2].Range["A2:J2"]);

                    if (dDayCur.Month >= 2 && dDayCur.Month <= 10)
                        workbook.Worksheets[2].Cells["A2"].Value = "Расписание экзаменнационной летней сессии " + (dDayCur.Year - 1).ToString() + "/" + dDayCur.Year.ToString() + " уч. года";
                    else
                        workbook.Worksheets[2].Cells["A2"].Value = "Расписание экзаменнационной зимней сессии " + dDayCur.Year.ToString() + "/" + (dDayCur.Year + 1).ToString() + " уч. года";

                    workbook.Worksheets[2].Rows["3"].Insert();
                    workbook.Worksheets[2].MergeCells(workbook.Worksheets[2].Range["A3:J3"]);
                    workbook.Worksheets[2].Rows["4"].Insert();
                    workbook.Worksheets[2].MergeCells(workbook.Worksheets[2].Range["A4:J4"]);
                    workbook.Worksheets[2].Cells["A4"].Value = "Кафедра \"Автоматизированные системы управления\"";
                    workbook.Worksheets[2].Rows["5"].Insert();
                    workbook.Worksheets[2].MergeCells(workbook.Worksheets[2].Range["A5:J5"]);
                    workbook.Worksheets[2].Range["A1:J6"].Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center;
                    workbook.Worksheets[2].Range["A1:J4"].Font.Size = 16;
                    workbook.Worksheets[2].Range["A1:J4"].Font.FontStyle = SpreadsheetFontStyle.BoldItalic;
                    workbook.Worksheets[2].Range["A1:J6"].Borders.SetAllBorders(Color.Black, BorderLineStyle.None);
                    workbook.Worksheets[2].Range["A6:J6"].Borders.SetAllBorders(Color.Black, BorderLineStyle.Thin);
                    workbook.Worksheets[2].Range["A6:J6"].Borders.SetOutsideBorders(Color.Black, BorderLineStyle.Thick);
                    workbook.Worksheets[2].ScrollToRow(0);

                    if (!UIButtons.UIButtonsEnabled(windowsUIButtonPanelExcel, new bool[] { false, false, false, false, true, false, false, false, false, true })) Console.WriteLine("Что то пошло не так!!!");

                    SplashScreenManager.CloseForm(false);

                    break;
                case "Очистить":
                    SplashScreenManager.ShowForm(typeof(Сессия.Forms.WaitFormUpdate));
                    SplashScreenManager.Default.SetWaitFormCaption("Ожидайте пожалуйста.");
                    SplashScreenManager.Default.SetWaitFormDescription("Очистка рабочей книги");

                    i = worksheets.Count - 1;

                    if (i > 0)
                        do
                        {
                            workbook.Worksheets.RemoveAt(i);
                            i--;
                        } while (i > 0);

                    workbook.Worksheets.Insert(0, "Лист1");
                    workbook.Worksheets.ActiveWorksheet = workbook.Worksheets[1];
                    workbook.Worksheets.RemoveAt(1);

                    if (sessionDBlDataSet.WorksheetData.Count() > 0)
                    {
                        if (!UIButtons.UIButtonsEnabled(windowsUIButtonPanelExcel, new bool[] { true, true, false, false, false, false, false, false, false, false })) Console.WriteLine("Что то пошло не так!!!");
                    }
                    else
                        if (!UIButtons.UIButtonsEnabled(windowsUIButtonPanelExcel, new bool[] { true, false, false, false, false, false, false, false, false, false })) Console.WriteLine("Что то пошло не так!!!");

                    SplashScreenManager.CloseForm(false);

                    break;
            }
        }
        #endregion

        #region Навигационные клавиши GridControl
        private void gridControlRooms_EmbeddedNavigator_ButtonClick(object sender, NavigatorButtonClickEventArgs e)
        {
            ColumnView view = (ColumnView)gridControlRooms.FocusedView;

            switch (((Control)sender).Parent.Name)
            {
                case "gridControlRooms":
                    view = (ColumnView)gridControlRooms.FocusedView;
                    break;
                case "gridControlHoliday":
                    view = (ColumnView)gridControlHoliday.FocusedView;
                    break;
            }

            switch (e.Button.ButtonType)
            {
                case NavigatorButtonType.Custom:
                    break;
                case NavigatorButtonType.First:
                    break;
                case NavigatorButtonType.PrevPage:
                    break;
                case NavigatorButtonType.Prev:
                    break;
                case NavigatorButtonType.Next:
                    break;
                case NavigatorButtonType.NextPage:
                    break;
                case NavigatorButtonType.Last:
                    break;
                case NavigatorButtonType.Append:
                    break;
                case NavigatorButtonType.Remove:
                    if (view == null || view.SelectedRowsCount == 0) break;
                    if (XtraMessageBox.Show(string.Format("Удалить отмеченную(-ые) {0} запись(-и)?", view.SelectedRowsCount), "Удаление данных", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    {
                        DataRow[] rows = new DataRow[view.SelectedRowsCount];
                        for (int iRow = 0; iRow < view.SelectedRowsCount; iRow++)
                            rows[iRow] = view.GetDataRow(view.GetSelectedRows()[iRow]);

                        view.BeginSort();
                        try
                        {
                            foreach (DataRow row in rows)
                                row.Delete();

                            this.tableAdapterManagerMain.UpdateAll(this.sessionDBlDataSet);
                        }
                        finally
                        {

                            view.EndSort();
                        }
                    }
                    else
                        e.Handled = true;

                    break;
                case NavigatorButtonType.Edit:
                    break;
                case NavigatorButtonType.EndEdit:
                    if (!(view.PostEditor() && view.UpdateCurrentRow())) return;
                    if (XtraMessageBox.Show("Записать в БД новые данные?", "Запись данных", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    {
                        this.Validate();

                        switch (((Control)sender).Parent.Name)
                        {
                            case "gridControlRooms":
                                this.roomsBindingSource.EndEdit();
                                break;
                            case "gridControlHoliday":
                                this.holidayBindingSource.EndEdit();
                                break;
                        }

                        this.tableAdapterManagerMain.UpdateAll(this.sessionDBlDataSet);
                    }

                    break;
                case NavigatorButtonType.CancelEdit:
                    break;
                default:
                    break;
            }
        }
        #endregion

        #region schedulerStorage1
        private void schedulerStorage1_AppointmentsChanged(object sender, PersistentObjectsEventArgs e)
        {
            CommitTask();
        }
        private void schedulerStorage1_AppointmentsDeleted(object sender, PersistentObjectsEventArgs e)
        {
            CommitTask();
        }
        private void schedulerStorage1_AppointmentsInserted(object sender, PersistentObjectsEventArgs e)
        {

            CommitTask();
            schedulerStorage1.SetAppointmentId(((Appointment)e.Objects[0]), id);
        }
        private void schedulerStorage1_AppointmentDependenciesChanged(object sender, PersistentObjectsEventArgs e)
        {
            CommitTaskDependency();
        }
        private void schedulerStorage1_AppointmentDependenciesDeleted(object sender, PersistentObjectsEventArgs e)
        {
            CommitTaskDependency();
        }
        private void schedulerStorage1_AppointmentDependenciesInserted(object sender, PersistentObjectsEventArgs e)
        {
            CommitTaskDependency();
        }
        void CommitTask()
        {
            appointmentsTableAdapter.Update(sessionDBlDataSet);
            this.sessionDBlDataSet.AcceptChanges();
        }
        void CommitTaskDependency()
        {
            taskDependenciesTableAdapter.Update(this.sessionDBlDataSet);
            this.sessionDBlDataSet.AcceptChanges();
        }
        #endregion

        #region #schedulerControlTimeTable
        private void schedulerControlTimeTable_ActiveViewChanged(object sender, EventArgs e)
        {
            this.splitContainerControlScheduler.SplitterPositionChanged -= new System.EventHandler(this.splitContainerControlScheduler_SplitterPositionChanged);
            bool isGanttView = schedulerControlTimeTable.ActiveViewType == SchedulerViewType.Gantt;

            try
            {
                splitContainerControlScheduler.SplitterPosition = (isGanttView) ? LastSplitContainerControlSplitterPosition : 0;
            }
            finally
            {
                this.splitContainerControlScheduler.SplitterPositionChanged += new System.EventHandler(this.splitContainerControlScheduler_SplitterPositionChanged);
            }
        }

        void schedulerControlTimeTable_InitNewAppointment(object sender, AppointmentEventArgs e)
        {
            if (e.Appointment.End.Date > e.Appointment.Start.Date)
                e.Appointment.End = new DateTime(e.Appointment.Start.Year, e.Appointment.Start.Month, e.Appointment.Start.Day, 19, 0, 0);
        }

        #region #tooltip_EmptySubject
        private void schedulerControlTimeTable_AppointmentViewInfoCustomizing(object sender, AppointmentViewInfoCustomizingEventArgs e)
        {
            if (e.ViewInfo.DisplayText == String.Empty)
                e.ViewInfo.ToolTipText = String.Format("Started at {0:g}", e.ViewInfo.Appointment.Start);
        }
        #endregion #tooltip_EmptySubject
        #endregion

        #region #EditAppointmentFormShowing
        private void schedulerControlTimeTable_EditAppointmentFormShowing(object sender, AppointmentFormEventArgs e)
        {
            CustomAppointmentForm form = new CustomAppointmentForm(sender as SchedulerControl, e.Appointment, e.OpenRecurrenceForm);
            try
            {
                e.DialogResult = form.ShowDialog();
                e.Handled = true;
            }
            finally
            {
                form.Dispose();
            }

            /*
                        XtraForm form;
                        // Create a form.
                        form = new AppointmentFormRibbonStyle(schedulerControlTimeTable, e.Appointment, e.OpenRecurrenceForm);
                        // Comply with restrictions.
                        ((AppointmentFormRibbonStyle)form).ReadOnly = e.ReadOnly;
                        form.LookAndFeel.ParentLookAndFeel = schedulerControlTimeTable.LookAndFeel;
                        e.DialogResult = form.ShowDialog(e.Parent);
                        e.Handled = true;
            */
        }
        #endregion

        #region splitContainerControlScheduler
        private void splitContainerControlScheduler_SplitterPositionChanged(object sender, EventArgs e)
        {
            LastSplitContainerControlSplitterPosition = splitContainerControlScheduler.SplitterPosition;
        }
        #endregion

        #region #ToolTipControllerBeforeShow
        private void toolTipControllerScheduler_BeforeShow(object sender, ToolTipControllerShowEventArgs e)
        {
            AppointmentViewInfo aptViewInfo;
            ToolTipController controller = (ToolTipController)sender;
            try
            {
                aptViewInfo = (AppointmentViewInfo)controller.ActiveObject;
            }
            catch
            {
                return;
            }

            if (aptViewInfo == null) return;

            if (toolTipControllerScheduler.ToolTipType == ToolTipType.Standard)
            {
                e.IconType = ToolTipIconType.Information;
                e.ToolTip = aptViewInfo.Description;
            }

            if (toolTipControllerScheduler.ToolTipType == ToolTipType.SuperTip)
            {
                SuperToolTip SuperTip = new SuperToolTip();
                SuperToolTipSetupArgs args = new SuperToolTipSetupArgs();
                args.Title.Text = aptViewInfo.Appointment.Location.ToString();
                args.Title.Font = new Font("Times New Roman", 14);
                args.Contents.Text = aptViewInfo.Description;
                args.Contents.Image = resImage;
                args.ShowFooterSeparator = true;
                args.Footer.Font = new Font("Comic Sans MS", 8);
                args.Footer.Text = aptViewInfo.Appointment.Subject.ToString();
                SuperTip.Setup(args);
                e.SuperTip = SuperTip;
            }
        }
        #endregion


        #region #Пользовательские функции и процедуры
        void GetDataSet()
        {
            // TODO: данная строка кода позволяет загрузить данные в таблицу "sessionDBlDataSet.TaskDependencies". При необходимости она может быть перемещена или удалена.
            this.taskDependenciesTableAdapter.Fill(this.sessionDBlDataSet.TaskDependencies);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "sessionDBlDataSet.Resources". При необходимости она может быть перемещена или удалена.
            this.resourcesTableAdapter.Fill(this.sessionDBlDataSet.Resources);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "sessionDBlDataSet.Appointments". При необходимости она может быть перемещена или удалена.
            this.appointmentsTableAdapter.Fill(this.sessionDBlDataSet.Appointments);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "sessionDBlDataSet1.Holiday". При необходимости она может быть перемещена или удалена.
            this.holidayTableAdapter.Fill(this.sessionDBlDataSet.Holiday);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "sessionDBlDataSet.Rooms". При необходимости она может быть перемещена или удалена.
            this.roomsTableAdapter.Fill(this.sessionDBlDataSet.Rooms);
            // Мероприятия
            this.schedulerStorage1.Appointments.Labels.Clear();
            this.schedulerStorage1.Appointments.Labels.Add(System.Drawing.SystemColors.Window, "Нет", "&Нет");
            this.schedulerStorage1.Appointments.Labels.Add(System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(194)))), ((int)(((byte)(190))))), "ГЭК", "&ГЭК");
            this.schedulerStorage1.Appointments.Labels.Add(System.Drawing.Color.FromArgb(((int)(((byte)(193)))), ((int)(((byte)(244)))), ((int)(((byte)(156))))), "Экзамен", "&Экзамен");
            this.schedulerStorage1.Appointments.Labels.Add(System.Drawing.Color.FromArgb(((int)(((byte)(244)))), ((int)(((byte)(206)))), ((int)(((byte)(147))))), "Доп. экзамен", "&Доп. экзамен");
            this.schedulerStorage1.Appointments.Labels.Add(System.Drawing.Color.FromArgb(((int)(((byte)(243)))), ((int)(((byte)(228)))), ((int)(((byte)(199))))), "Зачет", "&Зачет");
            this.schedulerStorage1.Appointments.Labels.Add(System.Drawing.Color.FromArgb(((int)(((byte)(199)))), ((int)(((byte)(244)))), ((int)(((byte)(255))))), "Доп. зачет", "До&п. зачет");
            this.schedulerStorage1.Appointments.Labels.Add(System.Drawing.Color.FromArgb(((int)(((byte)(168)))), ((int)(((byte)(213)))), ((int)(((byte)(255))))), "Консультация", "&Консультация");
            this.schedulerStorage1.Appointments.Labels.Add(System.Drawing.Color.FromArgb(((int)(((byte)(207)))), ((int)(((byte)(219)))), ((int)(((byte)(152))))), "Иное", "&Иное");
        }

        Form ShowResult(DataTable result)
        {
            Form newForm = new Form();
            newForm.Width = 600;
            newForm.Height = 300;

            DevExpress.XtraGrid.GridControl grid = new DevExpress.XtraGrid.GridControl();
            grid.Dock = DockStyle.Fill;
            grid.DataSource = result;

            newForm.Controls.Add(grid);
            grid.ForceInitialize();
            ((DevExpress.XtraGrid.Views.Grid.GridView)grid.FocusedView).OptionsView.ShowGroupPanel = false;

            newForm.ShowDialog(this);
            return newForm;
        }

        void exporter_CellValueConversionError(object sender, CellValueConversionErrorEventArgs e)
        {
            MessageBox.Show("Ошибка в ячейке " + e.Cell.GetReferenceA1());
            e.DataTableValue = null;
            e.Action = DataTableExporterAction.Continue;
        }
        #endregion
    }
}
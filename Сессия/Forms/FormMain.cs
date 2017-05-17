using DevExpress.Spreadsheet;
using DevExpress.Spreadsheet.Export;
using DevExpress.XtraBars.Docking2010;
using DevExpress.XtraEditors;
using DevExpress.XtraGrid.Views.Base;
using DevExpress.XtraScheduler;
using DevExpress.XtraSplashScreen;
using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Сессия.Classes;

namespace Сессия
{
    public partial class FormMain : DevExpress.XtraEditors.XtraForm
    {
        private int LastSplitContainerControlSplitterPosition;

        string[] connectionString = Utilities.GetConnectionStrings();
        int iCS = 0, id = 0;

        public FormMain()
        {
            InitializeComponent();

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

            //Fix the view type and splitter position.
            schedulerControlTimeTable.ActiveViewChanged += new EventHandler(schedulerControlTimeTable_ActiveViewChanged);
            this.splitContainerControlScheduler.SplitterPositionChanged += new System.EventHandler(this.splitContainerControlScheduler_SplitterPositionChanged);

            // Set the date to show existing appointments from the database.
            schedulerControlTimeTable.Start = DateTime.Today;
            schedulerControlTimeTable.ActiveViewType = DevExpress.XtraScheduler.SchedulerViewType.Timeline;
            schedulerControlTimeTable.InitNewAppointment += new AppointmentEventHandler(schedulerControlTimeTable_InitNewAppointment);

            #region  #scales
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

            #region #Adjustment
            //schedulerControlTimeTable.ActiveViewType = SchedulerViewType.Gantt;
            //schedulerControlTimeTable.ActiveViewType = SchedulerViewType.Timeline;
            schedulerControlTimeTable.GroupType = SchedulerGroupType.Resource;
            schedulerControlTimeTable.GanttView.CellsAutoHeightOptions.Enabled = true;
            // Hide unnecessary visual elements.
            schedulerControlTimeTable.GanttView.ShowResourceHeaders = false;
            schedulerControlTimeTable.GanttView.NavigationButtonVisibility = NavigationButtonVisibility.Never;
            // Disable user sorting in the Resource Tree (clicking the column will not change the sort order).
            колDescription.OptionsColumn.AllowSort = false;
            #endregion #Adjustment
            #endregion #scales
        }

        private void FormMain_Load(object sender, EventArgs e)
        {
            GetDataSet();

            for (int i = 0; i < connectionString.Length; i++)
            {
                string subString = "Сессия.Properties.Settings.SessionDBlConnectionString";

                if (connectionString[i].Trim() == subString)
                {
                    connectionString[i] = ConfigurationManager.ConnectionStrings[connectionString[i].Trim()].ConnectionString;
                    iCS = i;
                }
            }

            #region #CommitIdToDataSource
            schedulerStorage1.Appointments.CommitIdToDataSource = false;
            this.appointmentsTableAdapter.Adapter.RowUpdated += new SqlRowUpdatedEventHandler(appointmentsTableAdapter_RowUpdated);
            #endregion #CommitIdToDataSource

            if (!UIButtons.UIButtonsEnabled(windowsUIButtonPanelExcel, new bool[] { true, true, false, false, false, false, false, false, false })) Console.WriteLine("Что то пошло не так!!!");
        }

        #region Работа с БД
        private void GetDataSet()
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
        }

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
            bool bJob = false;

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

                            if (!UIButtons.UIButtonsEnabled(windowsUIButtonPanelExcel, new bool[] { false, false, false, true, true, false, false, false, true })) Console.WriteLine("Что то пошло не так!!!");
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
                                if (!UIButtons.UIButtonsEnabled(windowsUIButtonPanelExcel, new bool[] { true, false, false, false, false, false, false, false, false })) Console.WriteLine("Что то пошло не так!!!");

                            break;
                        }
                    }

                    workbook.LoadDocument(receivedBytes, DevExpress.Spreadsheet.DocumentFormat.OpenXml);

                    if (!UIButtons.UIButtonsEnabled(windowsUIButtonPanelExcel, new bool[] { false, false, false, true, true, false, false, false, true })) Console.WriteLine("Что то пошло не так!!!");

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
                    workbook.Worksheets[1].Cells["G1"].Value = "Status";
                    workbook.Worksheets[1].Cells["H1"].Value = "Label";
                    workbook.Worksheets[1].Cells["I1"].Value = "ResourceId";
                    workbook.Worksheets[1].Cells["J1"].Value = "PercentComplete";
                    workbook.Worksheets[1].Cells["K1"].Value = "TimeZoneId";

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
                            workbook.Worksheets[1].Cells[jRow, 4].Value = workbook.Worksheets[0].GetCellValue(0, iRow).ToString().Trim() + "; " +
                                                                          workbook.Worksheets[0].GetCellValue(2, iRow).ToString().Trim();
                            workbook.Worksheets[1].Cells[jRow, 5].Value = workbook.Worksheets[0].GetCellValue(4, iRow).ToString().Trim();
                            workbook.Worksheets[1].Cells[jRow, 6].Value = 2;
                            workbook.Worksheets[1].Cells[jRow, 7].Value = 3;

                            using (SqlConnection connection = new SqlConnection(connectionString[iCS]))
                            {
                                connection.Open();
                                SqlCommand command = connection.CreateCommand();
                                command.CommandText = string.Empty;
                                command.CommandText = "SELECT [Id] FROM Resources WHERE ([Description] Like N'" + workbook.Worksheets[0].GetCellValue(3, iRow).ToString().Trim() + "');";
                                workbook.Worksheets[1].Cells[jRow, 8].Value = Int32.Parse(Convert.ToString(command.ExecuteScalar()));
                            }

                            if (Convert.ToDateTime(workbook.Worksheets[1].GetCellValue(2, jRow).ToString().Trim()) < DateTime.Today)
                                workbook.Worksheets[1].Cells[jRow, 9].Value = 100;

                            workbook.Worksheets[1].Cells[jRow, 10].Value = "Russian Standard Time";
                            workbook.Worksheets[1].ScrollToRow(jRow);
                            iRow += 1;
                            jRow += 1;
                        }

                        iRow += 1;
                    } while (iRow <= usedRange.RowCount);

                    workbook.Worksheets[1].ScrollToRow(0);
                    workbook.Worksheets[1].Columns.AutoFit(0, 10);
                    bJob = true;

                    if (!UIButtons.UIButtonsEnabled(windowsUIButtonPanelExcel, new bool[] { false, false, false, false, true, false, true, false, true })) Console.WriteLine("Что то пошло не так!!!");

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

                    if (bJob)
                    {
                        if (!UIButtons.UIButtonsEnabled(windowsUIButtonPanelExcel, new bool[] { false, false, false, false, true, false, false, false, true })) Console.WriteLine("Что то пошло не так!!!");
                    }
                    else
                        if (!UIButtons.UIButtonsEnabled(windowsUIButtonPanelExcel, new bool[] { false, false, false, true, true, false, false, false, true })) Console.WriteLine("Что то пошло не так!!!");

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
                        string message = string.Format("Cannot update data in a database table.\n{0}", ex.Message);
                        MessageBox.Show(message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                    GetDataSet();

                    if (!UIButtons.UIButtonsEnabled(windowsUIButtonPanelExcel, new bool[] { false, false, false, false, true, false, false, false, true })) Console.WriteLine("Что то пошло не так!!!");

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
                    bJob = false;

                    if (!UIButtons.UIButtonsEnabled(windowsUIButtonPanelExcel, new bool[] { true, true, false, false, false, false, false, false, false })) Console.WriteLine("Что то пошло не так!!!");

                    SplashScreenManager.CloseForm(false);

                    break;
            }
        }

        // A custom converter that converts DateTime values to "Month-Year" text strings.
        class DateTimeToStringConverter : ICellValueToColumnTypeConverter
        {
            public bool SkipErrorValues { get; set; }
            public CellValue EmptyCellValue { get; set; }

            public ConversionResult Convert(Cell readOnlyCell, CellValue cellValue, Type dataColumnType, out object result)
            {
                result = DBNull.Value;
                ConversionResult converted = ConversionResult.Success;
                if (cellValue.IsEmpty)
                {
                    result = EmptyCellValue;
                    return converted;
                }
                if (cellValue.IsError)
                {
                    // You can return an error, subsequently the exporter throws an exception if the CellValueConversionError event is unhandled.
                    //return SkipErrorValues ? ConversionResult.Success : ConversionResult.Error;
                    result = "N/A";
                    return ConversionResult.Success;
                }
                result = String.Format("{0:MMMM-yyyy}", cellValue.DateTimeValue);
                return converted;
            }
        }

        void exporter_CellValueConversionError(object sender, CellValueConversionErrorEventArgs e)
        {
            MessageBox.Show("Error in cell " + e.Cell.GetReferenceA1());
            e.DataTableValue = null;
            e.Action = DataTableExporterAction.Continue;
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
        void CommitTask()
        {
            appointmentsTableAdapter.Update(sessionDBlDataSet);
            this.sessionDBlDataSet.AcceptChanges();
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
        void CommitTaskDependency()
        {
            taskDependenciesTableAdapter.Update(this.sessionDBlDataSet);
            this.sessionDBlDataSet.AcceptChanges();
        }
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
        #endregion

        #region splitContainerControlScheduler
        private void splitContainerControlScheduler_SplitterPositionChanged(object sender, EventArgs e)
        {
            LastSplitContainerControlSplitterPosition = splitContainerControlScheduler.SplitterPosition;
        }
        #endregion
    }
}
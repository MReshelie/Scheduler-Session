namespace Сессия
{
    partial class FormMain
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
            DevExpress.XtraEditors.TileItemElement tileItemElement1 = new DevExpress.XtraEditors.TileItemElement();
            DevExpress.Utils.SuperToolTip superToolTip1 = new DevExpress.Utils.SuperToolTip();
            DevExpress.Utils.ToolTipTitleItem toolTipTitleItem1 = new DevExpress.Utils.ToolTipTitleItem();
            DevExpress.Utils.ToolTipItem toolTipItem1 = new DevExpress.Utils.ToolTipItem();
            DevExpress.Utils.ToolTipSeparatorItem toolTipSeparatorItem1 = new DevExpress.Utils.ToolTipSeparatorItem();
            DevExpress.Utils.ToolTipTitleItem toolTipTitleItem2 = new DevExpress.Utils.ToolTipTitleItem();
            DevExpress.XtraEditors.TileItemElement tileItemElement2 = new DevExpress.XtraEditors.TileItemElement();
            DevExpress.Utils.SuperToolTip superToolTip2 = new DevExpress.Utils.SuperToolTip();
            DevExpress.Utils.ToolTipTitleItem toolTipTitleItem3 = new DevExpress.Utils.ToolTipTitleItem();
            DevExpress.Utils.ToolTipItem toolTipItem2 = new DevExpress.Utils.ToolTipItem();
            DevExpress.Utils.ToolTipSeparatorItem toolTipSeparatorItem2 = new DevExpress.Utils.ToolTipSeparatorItem();
            DevExpress.Utils.ToolTipTitleItem toolTipTitleItem4 = new DevExpress.Utils.ToolTipTitleItem();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormMain));
            this.tileBar = new DevExpress.XtraBars.Navigation.TileBar();
            this.tileBarGroupTables = new DevExpress.XtraBars.Navigation.TileBarGroup();
            this.eployeesTileBarItem = new DevExpress.XtraBars.Navigation.TileBarItem();
            this.customersTileBarItem = new DevExpress.XtraBars.Navigation.TileBarItem();
            this.navigationFrame = new DevExpress.XtraBars.Navigation.NavigationFrame();
            this.employeesNavigationPage = new DevExpress.XtraBars.Navigation.NavigationPage();
            this.layoutControl1 = new DevExpress.XtraLayout.LayoutControl();
            this.windowsUIButtonPanelExcel = new DevExpress.XtraBars.Docking2010.WindowsUIButtonPanel();
            this.spreadsheetControlExcel = new DevExpress.XtraSpreadsheet.SpreadsheetControl();
            this.layoutControlGroup1 = new DevExpress.XtraLayout.LayoutControlGroup();
            this.layoutControlItem1 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem2 = new DevExpress.XtraLayout.LayoutControlItem();
            this.employeesLabelControl = new DevExpress.XtraEditors.LabelControl();
            this.customersNavigationPage = new DevExpress.XtraBars.Navigation.NavigationPage();
            this.customersLabelControl = new DevExpress.XtraEditors.LabelControl();
            this.openFileDialogExcel = new System.Windows.Forms.OpenFileDialog();
            ((System.ComponentModel.ISupportInitialize)(this.navigationFrame)).BeginInit();
            this.navigationFrame.SuspendLayout();
            this.employeesNavigationPage.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControl1)).BeginInit();
            this.layoutControl1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlGroup1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem2)).BeginInit();
            this.customersNavigationPage.SuspendLayout();
            this.SuspendLayout();
            // 
            // tileBar
            // 
            this.tileBar.AllowDrag = false;
            this.tileBar.AllowGlyphSkinning = true;
            this.tileBar.AllowSelectedItem = true;
            this.tileBar.AppearanceGroupText.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(140)))), ((int)(((byte)(140)))), ((int)(((byte)(140)))));
            this.tileBar.AppearanceGroupText.Options.UseForeColor = true;
            this.tileBar.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(230)))), ((int)(((byte)(230)))), ((int)(((byte)(230)))));
            this.tileBar.Cursor = System.Windows.Forms.Cursors.Default;
            this.tileBar.Dock = System.Windows.Forms.DockStyle.Top;
            this.tileBar.DropDownButtonWidth = 30;
            this.tileBar.DropDownOptions.BeakColor = System.Drawing.Color.Empty;
            this.tileBar.Groups.Add(this.tileBarGroupTables);
            this.tileBar.IndentBetweenGroups = 10;
            this.tileBar.IndentBetweenItems = 10;
            this.tileBar.ItemPadding = new System.Windows.Forms.Padding(8, 6, 12, 6);
            this.tileBar.Location = new System.Drawing.Point(0, 0);
            this.tileBar.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.tileBar.MaxId = 4;
            this.tileBar.MaximumSize = new System.Drawing.Size(0, 110);
            this.tileBar.MinimumSize = new System.Drawing.Size(100, 110);
            this.tileBar.Name = "tileBar";
            this.tileBar.Padding = new System.Windows.Forms.Padding(29, 11, 29, 11);
            this.tileBar.ScrollMode = DevExpress.XtraEditors.TileControlScrollMode.None;
            this.tileBar.SelectedItem = this.eployeesTileBarItem;
            this.tileBar.SelectionBorderWidth = 2;
            this.tileBar.SelectionColorMode = DevExpress.XtraBars.Navigation.SelectionColorMode.UseItemBackColor;
            this.tileBar.ShowGroupText = false;
            this.tileBar.Size = new System.Drawing.Size(1223, 110);
            this.tileBar.TabIndex = 1;
            this.tileBar.Text = "tileBar";
            this.tileBar.WideTileWidth = 150;
            this.tileBar.SelectedItemChanged += new DevExpress.XtraEditors.TileItemClickEventHandler(this.tileBar_SelectedItemChanged);
            // 
            // tileBarGroupTables
            // 
            this.tileBarGroupTables.Items.Add(this.eployeesTileBarItem);
            this.tileBarGroupTables.Items.Add(this.customersTileBarItem);
            this.tileBarGroupTables.Name = "tileBarGroupTables";
            this.tileBarGroupTables.Text = "TABLES";
            // 
            // eployeesTileBarItem
            // 
            this.eployeesTileBarItem.AppearanceItem.Normal.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(115)))), ((int)(((byte)(196)))));
            this.eployeesTileBarItem.AppearanceItem.Normal.Options.UseBackColor = true;
            this.eployeesTileBarItem.DropDownOptions.BeakColor = System.Drawing.Color.Empty;
            tileItemElement1.ImageUri.Uri = "CustomizeGrid;Size32x32;GrayScaled";
            tileItemElement1.Text = "Импорт из Excel";
            this.eployeesTileBarItem.Elements.Add(tileItemElement1);
            this.eployeesTileBarItem.Name = "eployeesTileBarItem";
            superToolTip1.AllowHtmlText = DevExpress.Utils.DefaultBoolean.True;
            toolTipTitleItem1.Appearance.Image = global::Сессия.Properties.Resources.convert_32x32;
            toolTipTitleItem1.Appearance.Options.UseImage = true;
            toolTipTitleItem1.Image = global::Сессия.Properties.Resources.convert_32x32;
            toolTipTitleItem1.Text = "Импорт данных";
            toolTipItem1.LeftIndent = 6;
            toolTipItem1.Text = "Импорт в БД данных (полученны с официального сайта МАДИ), записанных в Excel-файл" +
    " в  формате xlsx.";
            toolTipTitleItem2.LeftIndent = 6;
            toolTipTitleItem2.Text = "Официальный сайт МАДИ (www.madi.ru), <href=http://tplan.madi.ru/task_manager.php?" +
    "task_id=10>Расписание экзаменов для кафедр</href>.";
            superToolTip1.Items.Add(toolTipTitleItem1);
            superToolTip1.Items.Add(toolTipItem1);
            superToolTip1.Items.Add(toolTipSeparatorItem1);
            superToolTip1.Items.Add(toolTipTitleItem2);
            this.eployeesTileBarItem.SuperTip = superToolTip1;
            // 
            // customersTileBarItem
            // 
            this.customersTileBarItem.DropDownOptions.BeakColor = System.Drawing.Color.Empty;
            tileItemElement2.ImageUri.Uri = "Customization;Size32x32;GrayScaled";
            tileItemElement2.Text = "Начальные установки";
            this.customersTileBarItem.Elements.Add(tileItemElement2);
            this.customersTileBarItem.Id = 2;
            this.customersTileBarItem.ItemSize = DevExpress.XtraBars.Navigation.TileBarItemSize.Wide;
            this.customersTileBarItem.Name = "customersTileBarItem";
            toolTipTitleItem3.Appearance.Image = global::Сессия.Properties.Resources.content_32x32;
            toolTipTitleItem3.Appearance.Options.UseImage = true;
            toolTipTitleItem3.Image = global::Сессия.Properties.Resources.content_32x32;
            toolTipTitleItem3.Text = "Начальные установки";
            toolTipItem2.LeftIndent = 6;
            toolTipItem2.Text = "Данные по списку аудиторий для учебного процесса";
            toolTipTitleItem4.LeftIndent = 6;
            toolTipTitleItem4.Text = "Общие справочники для функционирования системы расчетов";
            superToolTip2.Items.Add(toolTipTitleItem3);
            superToolTip2.Items.Add(toolTipItem2);
            superToolTip2.Items.Add(toolTipSeparatorItem2);
            superToolTip2.Items.Add(toolTipTitleItem4);
            this.customersTileBarItem.SuperTip = superToolTip2;
            // 
            // navigationFrame
            // 
            this.navigationFrame.Controls.Add(this.employeesNavigationPage);
            this.navigationFrame.Controls.Add(this.customersNavigationPage);
            this.navigationFrame.Dock = System.Windows.Forms.DockStyle.Fill;
            this.navigationFrame.Location = new System.Drawing.Point(0, 110);
            this.navigationFrame.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.navigationFrame.Name = "navigationFrame";
            this.navigationFrame.Pages.AddRange(new DevExpress.XtraBars.Navigation.NavigationPageBase[] {
            this.employeesNavigationPage,
            this.customersNavigationPage});
            this.navigationFrame.SelectedPage = this.employeesNavigationPage;
            this.navigationFrame.Size = new System.Drawing.Size(1223, 647);
            this.navigationFrame.TabIndex = 0;
            this.navigationFrame.Text = "navigationFrame";
            // 
            // employeesNavigationPage
            // 
            this.employeesNavigationPage.Caption = "employeesNavigationPage";
            this.employeesNavigationPage.Controls.Add(this.layoutControl1);
            this.employeesNavigationPage.Controls.Add(this.employeesLabelControl);
            this.employeesNavigationPage.Name = "employeesNavigationPage";
            this.employeesNavigationPage.Size = new System.Drawing.Size(1223, 647);
            // 
            // layoutControl1
            // 
            this.layoutControl1.Controls.Add(this.windowsUIButtonPanelExcel);
            this.layoutControl1.Controls.Add(this.spreadsheetControlExcel);
            this.layoutControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.layoutControl1.Location = new System.Drawing.Point(0, 0);
            this.layoutControl1.Name = "layoutControl1";
            this.layoutControl1.Root = this.layoutControlGroup1;
            this.layoutControl1.Size = new System.Drawing.Size(1223, 647);
            this.layoutControl1.TabIndex = 3;
            this.layoutControl1.Text = "layoutControl1";
            // 
            // windowsUIButtonPanelExcel
            // 
            this.windowsUIButtonPanelExcel.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(230)))), ((int)(((byte)(230)))), ((int)(((byte)(230)))));
            this.windowsUIButtonPanelExcel.Buttons.AddRange(new DevExpress.XtraEditors.ButtonPanel.IBaseButton[] {
            new DevExpress.XtraBars.Docking2010.WindowsUIButton("Открыть", "", DevExpress.XtraBars.Docking2010.ImageLocation.Default, DevExpress.XtraBars.Docking2010.ButtonStyle.PushButton, "", true, -1, true, null, true, false, true, null, "Ореn", -1, false, false),
            new DevExpress.XtraBars.Docking2010.WindowsUIButton("Обработать", "", DevExpress.XtraBars.Docking2010.ImageLocation.Default, DevExpress.XtraBars.Docking2010.ButtonStyle.PushButton, "", true, -1, true, null, true, false, true, null, "Work", -1, false, false),
            new DevExpress.XtraBars.Docking2010.WindowsUISeparator(),
            new DevExpress.XtraBars.Docking2010.WindowsUIButton("Очистить", "", DevExpress.XtraBars.Docking2010.ImageLocation.Default, DevExpress.XtraBars.Docking2010.ButtonStyle.PushButton, "", true, -1, true, null, true, false, true, null, "Clean", -1, false, false)});
            this.windowsUIButtonPanelExcel.ContentAlignment = System.Drawing.ContentAlignment.TopCenter;
            this.windowsUIButtonPanelExcel.Location = new System.Drawing.Point(12, 12);
            this.windowsUIButtonPanelExcel.Name = "windowsUIButtonPanelExcel";
            this.windowsUIButtonPanelExcel.Orientation = System.Windows.Forms.Orientation.Vertical;
            this.windowsUIButtonPanelExcel.Size = new System.Drawing.Size(65, 623);
            this.windowsUIButtonPanelExcel.TabIndex = 5;
            this.windowsUIButtonPanelExcel.Text = "windowsUIButtonPanel1";
            this.windowsUIButtonPanelExcel.ButtonClick += new DevExpress.XtraBars.Docking2010.ButtonEventHandler(this.windowsUIButtonPanelExcel_ButtonClick);
            // 
            // spreadsheetControlExcel
            // 
            this.spreadsheetControlExcel.Location = new System.Drawing.Point(81, 12);
            this.spreadsheetControlExcel.Name = "spreadsheetControlExcel";
            this.spreadsheetControlExcel.Size = new System.Drawing.Size(1130, 623);
            this.spreadsheetControlExcel.TabIndex = 4;
            this.spreadsheetControlExcel.Text = "spreadsheetControl1";
            // 
            // layoutControlGroup1
            // 
            this.layoutControlGroup1.EnableIndentsWithoutBorders = DevExpress.Utils.DefaultBoolean.True;
            this.layoutControlGroup1.GroupBordersVisible = false;
            this.layoutControlGroup1.Items.AddRange(new DevExpress.XtraLayout.BaseLayoutItem[] {
            this.layoutControlItem1,
            this.layoutControlItem2});
            this.layoutControlGroup1.Location = new System.Drawing.Point(0, 0);
            this.layoutControlGroup1.Name = "layoutControlGroup1";
            this.layoutControlGroup1.Size = new System.Drawing.Size(1223, 647);
            this.layoutControlGroup1.TextVisible = false;
            // 
            // layoutControlItem1
            // 
            this.layoutControlItem1.Control = this.spreadsheetControlExcel;
            this.layoutControlItem1.Location = new System.Drawing.Point(69, 0);
            this.layoutControlItem1.Name = "layoutControlItem1";
            this.layoutControlItem1.Size = new System.Drawing.Size(1134, 627);
            this.layoutControlItem1.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem1.TextVisible = false;
            // 
            // layoutControlItem2
            // 
            this.layoutControlItem2.Control = this.windowsUIButtonPanelExcel;
            this.layoutControlItem2.Location = new System.Drawing.Point(0, 0);
            this.layoutControlItem2.Name = "layoutControlItem2";
            this.layoutControlItem2.Size = new System.Drawing.Size(69, 627);
            this.layoutControlItem2.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem2.TextVisible = false;
            // 
            // employeesLabelControl
            // 
            this.employeesLabelControl.Appearance.Font = new System.Drawing.Font("Tahoma", 25.25F);
            this.employeesLabelControl.Appearance.ForeColor = System.Drawing.Color.Gray;
            this.employeesLabelControl.Appearance.Options.UseFont = true;
            this.employeesLabelControl.Appearance.Options.UseForeColor = true;
            this.employeesLabelControl.Appearance.Options.UseTextOptions = true;
            this.employeesLabelControl.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.employeesLabelControl.Appearance.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.employeesLabelControl.AutoSizeMode = DevExpress.XtraEditors.LabelAutoSizeMode.None;
            this.employeesLabelControl.Dock = System.Windows.Forms.DockStyle.Fill;
            this.employeesLabelControl.Location = new System.Drawing.Point(0, 0);
            this.employeesLabelControl.Name = "employeesLabelControl";
            this.employeesLabelControl.Size = new System.Drawing.Size(1223, 647);
            this.employeesLabelControl.TabIndex = 2;
            this.employeesLabelControl.Text = "Импорт из Excel";
            // 
            // customersNavigationPage
            // 
            this.customersNavigationPage.Caption = "customersNavigationPage";
            this.customersNavigationPage.Controls.Add(this.customersLabelControl);
            this.customersNavigationPage.Name = "customersNavigationPage";
            this.customersNavigationPage.Size = new System.Drawing.Size(1223, 647);
            // 
            // customersLabelControl
            // 
            this.customersLabelControl.Appearance.Font = new System.Drawing.Font("Tahoma", 25.25F);
            this.customersLabelControl.Appearance.ForeColor = System.Drawing.Color.Gray;
            this.customersLabelControl.Appearance.Options.UseFont = true;
            this.customersLabelControl.Appearance.Options.UseForeColor = true;
            this.customersLabelControl.Appearance.Options.UseTextOptions = true;
            this.customersLabelControl.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.customersLabelControl.Appearance.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.customersLabelControl.AutoSizeMode = DevExpress.XtraEditors.LabelAutoSizeMode.None;
            this.customersLabelControl.Dock = System.Windows.Forms.DockStyle.Fill;
            this.customersLabelControl.Location = new System.Drawing.Point(0, 0);
            this.customersLabelControl.Name = "customersLabelControl";
            this.customersLabelControl.Size = new System.Drawing.Size(1223, 647);
            this.customersLabelControl.TabIndex = 2;
            this.customersLabelControl.Text = "Справочники";
            // 
            // openFileDialogExcel
            // 
            this.openFileDialogExcel.FileName = "Сессия.xlsx";
            this.openFileDialogExcel.Filter = "Excel файл|*.xls;*.xlsx;*.xlsm";
            this.openFileDialogExcel.Title = "Поиск Excel-файла с расписанием";
            // 
            // FormMain
            // 
            this.Appearance.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(230)))), ((int)(((byte)(230)))), ((int)(((byte)(230)))));
            this.Appearance.Options.UseBackColor = true;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1223, 757);
            this.Controls.Add(this.navigationFrame);
            this.Controls.Add(this.tileBar);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "FormMain";
            this.Text = "Сессия, расписание мероприятий по аттестации студентов кафедрой";
            this.Load += new System.EventHandler(this.FormMain_Load);
            ((System.ComponentModel.ISupportInitialize)(this.navigationFrame)).EndInit();
            this.navigationFrame.ResumeLayout(false);
            this.employeesNavigationPage.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.layoutControl1)).EndInit();
            this.layoutControl1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlGroup1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem2)).EndInit();
            this.customersNavigationPage.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private DevExpress.XtraBars.Navigation.TileBar tileBar;
        private DevExpress.XtraBars.Navigation.NavigationFrame navigationFrame;
        private DevExpress.XtraBars.Navigation.TileBarGroup tileBarGroupTables;
        private DevExpress.XtraBars.Navigation.TileBarItem eployeesTileBarItem;
        private DevExpress.XtraBars.Navigation.TileBarItem customersTileBarItem;
        private DevExpress.XtraBars.Navigation.NavigationPage employeesNavigationPage;
        private DevExpress.XtraBars.Navigation.NavigationPage customersNavigationPage;
        private DevExpress.XtraEditors.LabelControl employeesLabelControl;
        private DevExpress.XtraEditors.LabelControl customersLabelControl;
        private DevExpress.XtraLayout.LayoutControl layoutControl1;
        private DevExpress.XtraBars.Docking2010.WindowsUIButtonPanel windowsUIButtonPanelExcel;
        private DevExpress.XtraSpreadsheet.SpreadsheetControl spreadsheetControlExcel;
        private DevExpress.XtraLayout.LayoutControlGroup layoutControlGroup1;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem1;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem2;
        private System.Windows.Forms.OpenFileDialog openFileDialogExcel;
    }
}
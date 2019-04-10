using System;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using System.Data.SqlClient;
using Excel = Microsoft.Office.Interop.Excel;
using System.Configuration;

namespace AccessList
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            Properties.Settings.Default.ExcelFilePath = this.textBoxOutput.Text;
            Properties.Settings.Default.Save();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.LoadSettings();
        }

        private void LoadSettings()
        {

            this.textBoxOutput.Text = Properties.Settings.Default.ExcelFilePath;

            Configuration config = ConfigurationManager.OpenExeConfiguration(Application.ExecutablePath);
            try
            {
                comboBoxEDSQLConnetction.Items.Add(config.AppSettings.Settings["EDProd"].Value);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error with config. file!" + Environment.NewLine + config.FilePath + Environment.NewLine + ex.Message);
                return;
            }
            
            comboBoxEDSQLConnetction.Items.Add(config.AppSettings.Settings["EDPred"].Value);
            comboBoxEDSQLConnetction.Items.Add(config.AppSettings.Settings["EDTest"].Value);
            comboBoxEDSQLConnetction.Items.Add(config.AppSettings.Settings["EDoro"].Value);
            comboBoxEDSQLConnetction.SelectedIndex = 0;

            comboBoxTRMSQLConnetction.Items.Add(config.AppSettings.Settings["TRMProd"].Value);
            comboBoxTRMSQLConnetction.Items.Add(config.AppSettings.Settings["TRMPred"].Value);
            comboBoxTRMSQLConnetction.Items.Add(config.AppSettings.Settings["TRMTest"].Value);            
            comboBoxTRMSQLConnetction.SelectedIndex = 0;

            comboBoxInvoicingSQLConnetction.Items.Add(config.AppSettings.Settings["INVProd"].Value);
            comboBoxInvoicingSQLConnetction.Items.Add(config.AppSettings.Settings["INVPred"].Value);
            comboBoxInvoicingSQLConnetction.Items.Add(config.AppSettings.Settings["INVTest"].Value);
            comboBoxInvoicingSQLConnetction.SelectedIndex = 0;
            
            int LTRolesCount = Convert.ToInt32(config.AppSettings.Settings["LTRoles_count"].Value);
            for (int i = 1; i <= LTRolesCount; i++)
            listBoxLT.Items.Add(config.AppSettings.Settings["LTRoles_" + i.ToString("D2")].Value);

            int LHRolesCount = Convert.ToInt32(config.AppSettings.Settings["LHRoles_count"].Value);
            for (int i = 1; i <= LHRolesCount; i++)
            listBoxLH.Items.Add(config.AppSettings.Settings["LHRoles_" + i.ToString("D2")].Value);

            int EDRolesCount = Convert.ToInt32(config.AppSettings.Settings["EDRoles_count"].Value);
            for (int i = 1; i <= EDRolesCount; i++)
            listBoxED.Items.Add(config.AppSettings.Settings["EDRoles_" + i.ToString("D2")].Value);

            this.Text = "Access list";

            checkBoxRoleED.Checked = true;
            checkBoxLT.Checked = true;
            checkBoxLH.Checked = true;

            dateTimePickerOldTree.Value = new DateTime(2019, 03, 30);
            dateTimePickerOldTree.Format = DateTimePickerFormat.Custom;
            dateTimePickerOldTree.CustomFormat = "yyyy.MM.dd";

            dateTimePickerNewTree.Value = DateTime.Today;
            dateTimePickerNewTree.Format = DateTimePickerFormat.Custom;
            dateTimePickerNewTree.CustomFormat = "yyyy.MM.dd";

            radioButtonNewTree.Checked = true;
            ToolboxSettings("ED");

            radioButtonFBUActiveAll.Checked = true;

            checkBoxActiveFBU.Checked = true;

            numericUpDownConTimeout.Value = 300;

            labelVersion.Text = "Version 0.1.1";

            /*
             * v0.1.1
             * 1. version added
             * 2. tree date updated
             * 3. access to parent FBU added
             * 4. move roles to appSettings.config
             */
        }

        private void buttonGenerateCurrentData_Click(object sender, EventArgs e)
        {
            if (checkBoxAudit.Checked == true)
            {
                GenerateAudit();
            }
            else
            {
                string SystemName = this.tabControlSystems.SelectedTab.Name;
                if (textBoxPrincipalAccess.Text != ""
                    || textBoxFBUAccess.Text != ""
                    || textBoxFBUPath.Text != ""
                    || (textBoxRoleAccess.Text != "" && SystemName != "FBUSearch" && SystemName != "FBUVersion")
                    || (textBoxHorizontal.Text != "" && SystemName == "FBUSearch")                    
                    )
                {
                    if (SystemName != "Settings")
                    {
                        string SQLRequest = GetReplacedSQLRequest(SystemName);
                        ExecuteSQLScript(SQLRequest, SystemName);
                        UpdateGridColor();
                    }
                }
            }
        }

        private string GetReplacedSQLRequest(string SystemName)
        {
            string Condition = "";
            string ConditionActive = "";
            string SQLRequestData = "";
            string FBUTable = "BusinessUnit";
            string FBUAncestorTable = "BusinessUnitAncestorLink";
            string DBName = "EnterpriseDirectories";

            switch (SystemName)
            {
                case "ED":
                    SQLRequestData = SQLQueriesTemplates.SQLEDAccess();                    
                    break;
                case "TRM":
                    SQLRequestData = SQLQueriesTemplates.SQLTRMAccess();
                    FBUTable = "FinancialBusinessUnit";
                    FBUAncestorTable = "FinancialBusinessUnitAncestorLink";
                    DBName = "TRMSys";
                    break;
                case "Invoicing":
                    SQLRequestData = SQLQueriesTemplates.SQLInvoicingAccess();
                    DBName = "Invoicing";
                    break;
                case "FBUManager":
                    SQLRequestData = SQLQueriesTemplates.SQLCurrentFBUManager();
                    break;
                case "FBUTree":
                    SQLRequestData = SQLQueriesTemplates.SQLTreeCompare();
                    break;
                case "FBUVersion":
                    SQLRequestData = SQLQueriesTemplates.SQLFBUVersionSearch();
                    break;
            }            

            //FBUSearch
            if (SystemName == "FBUTree")
            {
                string TreeType = "TreeOld";
                if (radioButtonOldTree.Checked == false)
                {
                    TreeType = "TreeNew";
                }
                if (textBoxFBUAccess.Text != "")
                {
                    Condition = Condition + TreeType + ".name like '%" + textBoxFBUAccess.Text + "%'";
                }
                if (textBoxFBUPath.Text != "")
                {
                    if (Condition != "")
                    {
                        Condition = Condition + " and ";
                    }
                    Condition = Condition + TreeType + ".Childs like '%" + textBoxFBUPath.Text + "%'";
                }
                if (textBoxHorizontal.Text != "")
                {
                    if (Condition != "")
                    {
                        Condition = Condition + " and ";
                    }
                    Condition = Condition + TreeType + ".Horizontal = '" + textBoxHorizontal.Text + "'";
                }
                //active
                if (radioButtonFBUActiveYes.Checked == true)
                {
                    if (Condition != "")
                    {
                        Condition = Condition + " and ";
                    }
                    Condition = Condition + "TreeNew.active = 1";
                }
                if (radioButtonFBUActiveNo.Checked == true)
                {
                    if (Condition != "")
                    {
                        Condition = Condition + " and ";
                    }
                    Condition = Condition + "TreeNew.active = 0";
                }
                SQLRequestData = SQLRequestData.Replace("#Condition", Condition);
                SQLRequestData = SQLRequestData.Replace("#OldTreeDate", dateTimePickerOldTree.Value.ToString("yyyyMMdd"));
                SQLRequestData = SQLRequestData.Replace("#NewTreeDate", dateTimePickerNewTree.Value.ToString("yyyyMMdd"));

                return SQLRequestData;
            }

            //FBUVersion
            if (textBoxFBUAccess.Text != "" && SystemName == "FBUVersion")
            {
                Condition = textBoxFBUAccess.Text;

                SQLRequestData = SQLRequestData.Replace("#Condition", Condition);
                return SQLRequestData;
            }           

            //all other
            if (textBoxPrincipalAccess.Text != "")
            {
                if (SystemName == "ED" || SystemName == "TRM" || SystemName == "Invoicing")
                {
                    Condition = "PermissionsAll.";
                }
                Condition = Condition + "Login like '%" + textBoxPrincipalAccess.Text + "%'";
            }
            
            if (textBoxFBUAccess.Text != "")
            {
                if (Condition != "")
                {
                    Condition = Condition + " and ";
                }
                if (checkBoxParentFBU.Checked)
                {
                    string SQLFBUParentFilter = SQLQueriesTemplates.SQLFBUParentFilter();
                    SQLFBUParentFilter = SQLFBUParentFilter.Replace("#ConditionFBU", textBoxFBUAccess.Text);
                    SQLFBUParentFilter = SQLFBUParentFilter.Replace("#BusinessUnitTableName", FBUTable);
                    SQLFBUParentFilter = SQLFBUParentFilter.Replace("#BusinessUnitAncestorTableName", FBUAncestorTable);
                    SQLFBUParentFilter = SQLFBUParentFilter.Replace("#DBName", DBName);
                    
                    Condition = Condition + "BU in (" + SQLFBUParentFilter + ")";
                }
                else {
                    Condition = Condition + "BU = '" + textBoxFBUAccess.Text + "'";
                }                
            }
            if (textBoxFBUPath.Text != "")
            {
                if (Condition != "")
                {
                    Condition = Condition + " and ";
                }
                Condition = Condition + "BU_path like '%" + textBoxFBUPath.Text + "%'";
            }
            if (textBoxRoleAccess.Text != "")
            {
                if (Condition != "")
                {
                    Condition = Condition + " and ";
                }
                //Condition = Condition + "Role like '%" + textBoxRoleAccess.Text + "%'";
                Condition = Condition + "Role = '" + textBoxRoleAccess.Text + "'";
            }
            if (checkBoxActiveFBU.Checked == true && SystemName != "FBUManager")
            {
                ConditionActive = ConditionActive + " and BusinessUnit.active = 1";
            }

            if (SystemName == "ED")
            {
                //собираем список ролей
                string LTRoles = GetRoleList(listBoxLT);
                string LHRoles = GetRoleList(listBoxLH);
                string EDRoles = GetRoleList(listBoxED);

                //Список ролей собрали, теперь собираем запрос
                if (checkBoxRoleED.Checked == false || checkBoxLT.Checked == false || checkBoxLH.Checked == false)
                {
                    if (checkBoxLT.Checked == false)
                    { Condition = Condition + @" and Role not in (" + LTRoles + ")"; }
                    //else {   Condition = Condition + @" and Role in (" + LTRoles + ")";  }

                    if (checkBoxLH.Checked == false)
                    { Condition = Condition + @" and Role not in (" + LHRoles + ")"; }
                    //else {   Condition = Condition + @" and Role in (" + LHRoles + ")";  }

                    if (checkBoxRoleED.Checked == false)
                    { Condition = Condition + @" and Role not in (" + EDRoles + ")"; }
                    //else {   Condition = Condition + @" and Role in (" + EDRoles + ")";  }
                }
            }

            SQLRequestData = SQLRequestData.Replace("#ConditionActiveFBU", ConditionActive);
            SQLRequestData = SQLRequestData.Replace("#Condition", Condition);

            return SQLRequestData;
        }

        private string GetRoleList(ListBox listBoxRoles)
        {
            string ListRoles = "";
            int listBoxCount = listBoxRoles.Items.Count;
            for (int i = 0; i < listBoxCount; i++)
            {
                if (i == 0)
                {
                    ListRoles = ListRoles + "'" + listBoxRoles.Items[i] + "'";
                }
                else
                {
                    ListRoles = ListRoles + ",'" + listBoxRoles.Items[i] + "'";
                }
            }

            return ListRoles;
        }

        private void GenerateAudit()
        {
            string SystemName = this.tabControlSystems.SelectedTab.Name;

            if (textBoxPrincipalAccess.Text != "" || ((textBoxPrincipalAccess.Text != "" || textBoxFBUAccess.Text != "") && SystemName == "FBUManager"))
            {
                string DatabaseToUse = "";
                string BusinessUnitTableName = "";
                string SQLRequest = "";
                switch (SystemName)
                {
                    case "ED":
                        DatabaseToUse = "EnterpriseDirectories";
                        BusinessUnitTableName = "BusinessUnit";
                        break;
                    case "TRM":
                        DatabaseToUse = "TRMSys";
                        BusinessUnitTableName = "[FinancialBusinessUnit] as BusinessUnit";
                        break;
                    case "Invoicing":
                        DatabaseToUse = "Invoicing";
                        BusinessUnitTableName = "BusinessUnit";
                        break;
                    case "FBUManager":
                        DatabaseToUse = "EnterpriseDirectories";
                        BusinessUnitTableName = "BusinessUnit";
                        break;
                }
                if (SystemName == "ED" || SystemName == "TRM" || SystemName == "Invoicing")
                {
                    SQLRequest = SQLQueriesTemplates.SQLAccessAudit();
                    SQLRequest = SQLRequest.Replace("#DatabaseToUse", DatabaseToUse);
                    SQLRequest = SQLRequest.Replace("#BusinessUnitTableName", BusinessUnitTableName);
                    SQLRequest = SQLRequest.Replace("#Condition", textBoxPrincipalAccess.Text);
                    
                    SQLRequest = SQLRequest.Replace("[Authorization].[dbo]", " [auth]");
                    SQLRequest = SQLRequest.Replace("[AuthorizationAudit].[dbo]", " [authAudit]");
                    SQLRequest = SQLRequest.Replace("[dbo]", " [app]");
                    
                }
                else
                {
                    string Condition = "";
                    if (textBoxPrincipalAccess.Text != "")
                    {
                        Condition = Condition + "Employee.login = '" + textBoxPrincipalAccess.Text + "'";
                    }
                    if (textBoxFBUAccess.Text != "")
                    {
                        if (Condition != "")
                        {
                            Condition = Condition + " and ";
                        }
                        Condition = Condition + "businessUnit.name = '" + textBoxFBUAccess.Text + "'";
                    }
                    SQLRequest = SQLQueriesTemplates.SQLFBUManagerAudit();
                    SQLRequest = SQLRequest.Replace("#Condition", Condition);
                }

                ExecuteSQLScript(SQLRequest, SystemName);
            }
        }

        private void UpdateGridColor()
        {
            foreach (DataGridViewRow row in dataGridViewFBUSearch.Rows)
            {
                if (row.Cells["FBU_New"].Value.ToString() != row.Cells["FBU_Old"].Value.ToString())
                {
                    row.Cells["FBU_New"].Style.ForeColor = Color.Red;
                }
                if (row.Cells["FBUParent_New"].Value.ToString() != row.Cells["FBUParent_Old"].Value.ToString())
                {
                    row.Cells["FBUParent_New"].Style.ForeColor = Color.Red;
                }
                if (row.Cells["FBUType_New"].Value.ToString() != row.Cells["FBUType_Old"].Value.ToString())
                {
                    row.Cells["FBUType_New"].Style.ForeColor = Color.Red;
                }
                if (row.Cells["FBU_path_New"].Value.ToString() != row.Cells["FBU_path_Old"].Value.ToString())
                {
                    row.Cells["FBU_path_New"].Style.ForeColor = Color.Red;
                }
                if (row.Cells["Horizontal_New"].Value.ToString() != row.Cells["Horizontal_Old"].Value.ToString())
                {
                    row.Cells["Horizontal_New"].Style.ForeColor = Color.Red;
                }
                if (row.Cells["FBU_Old"].Value.ToString() == "" && row.Cells["FBUStatus_New"].Value.ToString() == "Overdue")
                {
                    row.Cells["FBU_New"].Style.ForeColor = Color.Green;
                    row.Cells["FBUParent_New"].Style.ForeColor = Color.Green;
                    row.Cells["FBUType_New"].Style.ForeColor = Color.Green;
                    row.Cells["FBU_path_New"].Style.ForeColor = Color.Green;
                    row.Cells["Horizontal_New"].Style.ForeColor = Color.Green;
                }
            }
        }

        private void ExecuteSQLScript(string SQLRequest, string SystemName)
        {
            string DataSource = "";
            string InitialCatalog = "";
            DataGridView DataGridViewToUpdate = new DataGridView();
            switch (SystemName)
            {
                case "ED":
                    DataSource = "" + comboBoxEDSQLConnetction.SelectedItem;
                    InitialCatalog = "EnterpriseDirectories";
                    DataGridViewToUpdate = dataGridViewED;
                    break;
                case "TRM":
                    DataSource = "" + comboBoxTRMSQLConnetction.SelectedItem;
                    InitialCatalog = "TRMSys";
                    DataGridViewToUpdate = dataGridViewTRM;
                    break;
                case "Invoicing":
                    DataSource = "" + comboBoxInvoicingSQLConnetction.SelectedItem;
                    InitialCatalog = "Invoicing";
                    DataGridViewToUpdate = dataGridViewInvoicing;
                    break;
                case "FBUManager":
                    DataSource = "" + comboBoxEDSQLConnetction.SelectedItem;
                    InitialCatalog = "EnterpriseDirectories";
                    DataGridViewToUpdate = dataGridViewFBUManager;
                    break;
                case "FBUTree":
                    DataSource = "" + comboBoxEDSQLConnetction.SelectedItem;
                    InitialCatalog = "EnterpriseDirectories";
                    DataGridViewToUpdate = dataGridViewFBUSearch;
                    break;
                case "FBUVersion":
                    DataSource = "" + comboBoxEDSQLConnetction.SelectedItem;
                    InitialCatalog = "EnterpriseDirectories";
                    DataGridViewToUpdate = dataGridViewFBUVersion;
                    break;
            }            
            DataSet ds = new DataSet();
            SqlConnection con = new SqlConnection("Data Source=" + DataSource + ";Initial Catalog=" + InitialCatalog + ";Integrated Security=True");
            SqlDataAdapter da = new SqlDataAdapter();
            da.SelectCommand = new SqlCommand(SQLRequest, con);
            da.SelectCommand.CommandTimeout = Convert.ToInt32(numericUpDownConTimeout.Value);
            con.Open();
            da.Fill(ds, "PermissionsAll");
            DataGridViewToUpdate.DataSource = ds.Tables[0];
            da.Dispose();
            con.Dispose();
            ds.Dispose();

            DataGridViewToUpdate.AutoResizeColumns();
            this.tabControlSystems.SelectedTab.Text = this.tabControlSystems.SelectedTab.Name + " (" + DataGridViewToUpdate.RowCount + ")";
        }

        private void buttonFinReportToFile_Click(object sender, EventArgs e)
        {
            SaveDataToExcel(this.textBoxOutput.Text);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SaveAllDataToExcel(this.textBoxOutput.Text);
        }

        //Results to excel        
        private void SaveDataToExcel(string PathToOutput)
        {            
            DataGridView DataGridIOutput = new DataGridView();
            string FileName = this.tabControlSystems.SelectedTab.Name;
            string localDateStr = DateTime.Now.ToString().Replace(":", "").Replace(".", "");
            string PathFull = PathToOutput + "" + FileName + "_" + localDateStr + ".xlsx";
            switch (FileName)
            {
                case "ED":
                    DataGridIOutput = dataGridViewED;
                    break;
                case "TRM":
                    DataGridIOutput = dataGridViewTRM;
                    break;
                case "Invoicing":
                    DataGridIOutput = dataGridViewInvoicing;
                    break;
                case "FBUManager":
                    DataGridIOutput = dataGridViewFBUManager;
                    break;
                case "FBUTree":
                    DataGridIOutput = dataGridViewFBUSearch;
                    break;
                case "FBUVersion":
                    DataGridIOutput = dataGridViewFBUVersion;
                    break;
            }
            /*
            Excel.Application excelapp = new Excel.Application();
            Excel.Workbook workbook = excelapp.Workbooks.Add();
            Excel.Worksheet worksheet = workbook.Worksheets.Add();
            foreach (Excel.Worksheet ws in workbook.Worksheets)
            {
                if (worksheet != ws)
                {
                    ws.Delete();
                }
            }

            worksheet.Name = this.tabControlSystems.SelectedTab.Name;

            FillWorkSheet(DataGridIOutput, worksheet);
            excelapp.AlertBeforeOverwriting = false;
            workbook.SaveAs(PathFull);
            excelapp.Quit();

            MessageBox.Show("Save completed. " + PathFull);
            */

            ExcelData ExcelDataNew = new ExcelData();
            ExcelDataNew.ExportDataSet((DataTable)DataGridIOutput.DataSource, PathFull, FileName, true);

            Excel.Application ObjExcel = new Excel.Application();
            Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(PathFull, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);            
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];
            for (int j = 1; j < 10; j++)
            {
                ObjWorkSheet.Rows[1].Columns[j].EntireRow.Font.Bold = true;
                ObjWorkSheet.Rows[1].Columns[j].EntireColumn.AutoFit();            
            }
            ObjExcel.AlertBeforeOverwriting = false;
            ObjWorkBook.Save();
            ObjExcel.Quit();

            MessageBox.Show("Save completed. " + PathFull);
        }

        private void SaveAllDataToExcel(string PathToOutput)
        {
            string FileName = "AllAccess_";
            string localDateStr = DateTime.Now.ToString().Replace(":", "").Replace(".", "");
            string PathFull = PathToOutput + "" + FileName + "_" + localDateStr + ".xlsx";

            /*
            Excel.Application excelapp = new Excel.Application();
            Excel.Workbook workbook = excelapp.Workbooks.Add();
            Excel.Worksheet worksheet = workbook.ActiveSheet;
            foreach (Excel.Worksheet ws in workbook.Worksheets)
            {
                if (worksheet != ws)
                {
                    ws.Delete();
                }
            }

            if (dataGridViewFBUSearch.RowCount > 0)
            {
                Excel.Worksheet worksheet5 = workbook.Worksheets.Add();
                worksheet5.Name = "FBUTree";
                FillWorkSheet(dataGridViewFBUSearch, worksheet5);
            }
            if (dataGridViewFBUManager.RowCount > 0)
            {
                Excel.Worksheet worksheet4 = workbook.Worksheets.Add();
                worksheet4.Name = "FBUManager";
                FillWorkSheet(dataGridViewFBUManager, worksheet4);
            }
            if (dataGridViewInvoicing.RowCount > 0)
            {
                Excel.Worksheet worksheet3 = workbook.Worksheets.Add();
                worksheet3.Name = "Invoicing";
                FillWorkSheet(dataGridViewInvoicing, worksheet3);
            }
            if (dataGridViewTRM.RowCount > 0)
            {
                Excel.Worksheet worksheet2 = workbook.Worksheets.Add();
                worksheet2.Name = "TRM";
                FillWorkSheet(dataGridViewTRM, worksheet2);
            }
            if (dataGridViewED.RowCount > 0)
            {
                Excel.Worksheet worksheet1 = workbook.Worksheets.Add();
                worksheet1.Name = "ED";
                FillWorkSheet(dataGridViewED, worksheet1);
            }
            if (workbook.Worksheets.Count > 1)
            {
                worksheet.Delete();
            }

            string localDateStr = DateTime.Now.ToString().Replace(":", "").Replace(".", "");
            excelapp.AlertBeforeOverwriting = false;
            workbook.SaveAs(PathToOutput + "AllAccess_" + localDateStr + ".xlsx");
            excelapp.Quit();

            MessageBox.Show("Save completed. " + PathToOutput + "\\AllAccess_" + localDateStr + ".xlsx");
            */

            Int32 SheetsCount = 0;
            ExcelData ExcelDataNew = new ExcelData();
            if (dataGridViewED.RowCount > 0)
            {
                ExcelDataNew.ExportDataSet((DataTable)dataGridViewED.DataSource, PathFull, "ED", true);
                SheetsCount++;
            }
            if (dataGridViewTRM.RowCount > 0)
            {
                ExcelDataNew.ExportDataSet((DataTable)dataGridViewTRM.DataSource, PathFull, "TRM", false);
                SheetsCount++;
            }            
            if (dataGridViewInvoicing.RowCount > 0)
            {
                ExcelDataNew.ExportDataSet((DataTable)dataGridViewInvoicing.DataSource, PathFull, "Invoicing", false);
                SheetsCount++;
            }
            if (dataGridViewFBUManager.RowCount > 0)
            {
                ExcelDataNew.ExportDataSet((DataTable)dataGridViewFBUManager.DataSource, PathFull, "FBUManager", false);
                SheetsCount++;
            }
            if (dataGridViewFBUSearch.RowCount > 0)
            {
                ExcelDataNew.ExportDataSet((DataTable)dataGridViewFBUSearch.DataSource, PathFull, "FBUTree", false);
                SheetsCount++;
            }

            Excel.Application ObjExcel = new Excel.Application();
            Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(PathFull, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            for (Int32 i = 1; i <= SheetsCount; i++)
            {
                Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[i];
                for (int j = 1; j < 10; j++)
                {
                    ObjWorkSheet.Rows[1].Columns[j].EntireRow.Font.Bold = true;
                    ObjWorkSheet.Rows[1].Columns[j].EntireColumn.AutoFit();
                }
            }
            ObjExcel.AlertBeforeOverwriting = false;
            ObjWorkBook.Save();
            ObjExcel.Quit();

            MessageBox.Show("Save completed. " + PathFull);
        }

        private void FillWorkSheet(DataGridView DataGridIOutput, Excel.Worksheet worksheet)
        {
            //header
            for (int j = 1; j < DataGridIOutput.ColumnCount + 1; j++)
            {
                worksheet.Rows[1].Columns[j] = DataGridIOutput.Columns[j - 1].HeaderText;
                //worksheet.Rows[1].Columns[j].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                worksheet.Rows[1].Columns[j].EntireRow.Font.Bold = true;
                worksheet.Rows[1].Columns[j].EntireColumn.AutoFit();
            }

            //body            
            for (int i = 1; i < DataGridIOutput.RowCount + 1; i++)
            {
                for (int j = 1; j < DataGridIOutput.ColumnCount + 1; j++)
                {
                    worksheet.Rows[i + 1].Columns[j] = DataGridIOutput.Rows[i - 1].Cells[j - 1].Value.ToString();
                }
            }
            for (int j = 1; j < DataGridIOutput.ColumnCount + 1; j++)
            {
                worksheet.Columns[j].EntireColumn.AutoFit();
            }
        }
        

        private void buttonFinReportOutput_Click(object sender, EventArgs e)
        {
            DirectoryOpen_Click(this.textBoxOutput.Name);
        }

        //open directiory
        private void DirectoryOpen_Click(string PathName)
        {
            Control pathString = this.Controls.Find(PathName, true).First();
            using (var fbd = new FolderBrowserDialog())
            {
                DialogResult result = fbd.ShowDialog();

                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    pathString.Text = fbd.SelectedPath;
                }
            }
        }
        
        /////////////////////////////////////////////////////////
        //Form elements visiblity
        private void checkBoxAudit_CheckedChanged(object sender, EventArgs e)
        {
            SetAuditChecked();
        }

        private void SetAuditChecked()
        {
            if (checkBoxAudit.Checked == true)
            {
                textBoxFBUAccess.Enabled = false;
                textBoxFBUPath.Enabled = false;
                textBoxRoleAccess.Enabled = false;
                labelFBU.Enabled = false;
                labelFBUPath.Enabled = false;
                labelRole.Enabled = false;
                if (this.tabControlSystems.SelectedTab.Name == "FBUManager")
                {
                    textBoxFBUAccess.Enabled = true;
                    labelFBU.Enabled = true;
                }
            }
            else
            {
                textBoxFBUAccess.Enabled = true;
                textBoxFBUPath.Enabled = true;
                textBoxRoleAccess.Enabled = true;
                labelFBU.Enabled = true;
                labelFBUPath.Enabled = true;
                labelRole.Enabled = true;
                buttonClearFBUPath.Enabled = true;
            }
        }

        private void tabControlSystems_Selected(object sender, TabControlEventArgs e)
        {
            string TabPageName = e.TabPage.Name;
            ToolboxSettings(TabPageName);
        }

        private void ToolboxSettings(string TabPageName)
        {
            if (TabPageName == "FBUTree" || TabPageName == "FBUVersion")
            {
                textBoxPrincipalAccess.Enabled = false;
                labelPrincipal.Enabled = false;
                buttonClearPrincipal.Enabled = false;

                textBoxRoleAccess.Visible = false;
                labelRole.Visible = false;
                buttonClearRoleAccess.Visible = false;

                textBoxHorizontal.Visible = true;
                labelHorizontal.Visible = true;
                buttonClearHorizontal.Visible = true;
                textBoxHorizontal.Enabled = true;
                labelHorizontal.Enabled = true;
                buttonClearHorizontal.Enabled = true;

                textBoxFBUPath.Enabled = true;
                labelFBUPath.Enabled = true;
                buttonClearFBUPath.Enabled = true;
                if (TabPageName == "FBUVersion")
                {
                    textBoxFBUPath.Enabled = false;
                    labelFBUPath.Enabled = false;
                    buttonClearFBUPath.Enabled = false;

                    textBoxHorizontal.Enabled = false;
                    labelHorizontal.Enabled = false;
                    buttonClearHorizontal.Enabled = false;
                }
            }
            else
            {
                textBoxPrincipalAccess.Enabled = true;
                labelPrincipal.Enabled = true;
                buttonClearPrincipal.Enabled = true;

                textBoxRoleAccess.Visible = true;
                labelRole.Visible = true;
                buttonClearRoleAccess.Visible = true;

                textBoxFBUPath.Enabled = true;
                labelFBUPath.Enabled = true;
                buttonClearFBUPath.Enabled = true;

                textBoxHorizontal.Visible = false;
                labelHorizontal.Visible = false;
                buttonClearHorizontal.Visible = false;
                textBoxHorizontal.Enabled = true;
                labelHorizontal.Enabled = true;
                buttonClearHorizontal.Enabled = true;
            }

            SetAuditChecked();

            if (TabPageName == "ED" || TabPageName == "TRM" || TabPageName == "Invoicing" || TabPageName == "FBUManager" || TabPageName == "Settings")
            {
                checkBoxAudit.Enabled = true;
                checkBoxParentFBU.Enabled = true;
            }
            else
            {
                checkBoxAudit.Checked = false;
                checkBoxAudit.Enabled = false;
                checkBoxParentFBU.Checked = false;
                checkBoxParentFBU.Enabled = false;
            }

            if (TabPageName == "Settings")
            {
                buttonGenerateCurrentData.Enabled = false;
            }
            else
            {
                buttonGenerateCurrentData.Enabled = true;
            }
        }

        ////////////////////////////////////////////////
        //Clear textbox buttons
        private void buttonClearPrincipal_Click(object sender, EventArgs e)
        {
            textBoxPrincipalAccess.Text = "";
        }

        private void buttonClearFBU_Click(object sender, EventArgs e)
        {
            textBoxFBUAccess.Text = "";
        }

        private void buttonClearFBUPath_Click(object sender, EventArgs e)
        {
            textBoxFBUPath.Text = "";
        }

        private void buttonClearHorizontal_Click(object sender, EventArgs e)
        {
            textBoxHorizontal.Text = "";
        }

        private void buttonClearRoleAccess_Click(object sender, EventArgs e)
        {
            textBoxRoleAccess.Text = "";
        }
        
        ////////////////////////////////////////////////
        //Clear grid buttons
        private void buttonClearDataED_Click(object sender, EventArgs e)
        {
            ClearGridData(dataGridViewED);
        }

        private void buttonClearDataTRM_Click(object sender, EventArgs e)
        {
            ClearGridData(dataGridViewTRM);
        }

        private void buttonClearDataInvoicing_Click(object sender, EventArgs e)
        {
            ClearGridData(dataGridViewInvoicing);
        }

        private void buttonClearDataFBUManager_Click(object sender, EventArgs e)
        {
            ClearGridData(dataGridViewFBUManager);
        }

        private void buttonClearFBUVersion_Click(object sender, EventArgs e)
        {
            ClearGridData(dataGridViewFBUVersion);
        }

        private void buttonClearFBUTree_Click(object sender, EventArgs e)
        {
            ClearGridData(dataGridViewFBUSearch);
        }

        private void ClearGridData(DataGridView DataGridViewToUpdate)
        {
            if (DataGridViewToUpdate.DataSource == null)
            {
                DataGridViewToUpdate.Rows.Clear();
                DataGridViewToUpdate.Columns.Clear();
            }
            else
            {
                DataGridViewToUpdate.DataSource = null;
            }
            this.tabControlSystems.SelectedTab.Text = this.tabControlSystems.SelectedTab.Name;
        }
    }
}

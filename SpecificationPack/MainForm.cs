using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace SpecificationPack
{
    public partial class MainForm : Form
    {
        private List<Unit> Units;
        private List<GroupUnit> GroupUnits;
        private List<string> Except;
        private Excel.Application excel;
        private Excel.Window excelWindow;
        

        struct Unit
        {
            public string Group;
            public string Code;
            public string Name;
            public string Manufacture;
            public CupBoard[] cupBoard;
            public string Measure;
            public Color errorColor;

            public double Count;
            public string FileName;
        }

        public struct CupBoard
        {
            public double Num;
            public string fileName;
        }

        public struct GroupUnit
        {
            public string Group;
            public string Code;
            public string Name;
        }

        public MainForm()
        {
            InitializeComponent();
        }

        private void addSpecBtn_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Multiselect = true;
            ofd.Filter = "(*.xlsx); (*.xls)|*.xlsx; *.xls";
            if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                foreach (string fileName in ofd.FileNames)
                    specListBox.Items.Add(fileName);
            }
        }

        private void specListBox_DragDrop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop) && e.Effect == DragDropEffects.Move)
            {
                string[] objects = (string[])e.Data.GetData(DataFormats.FileDrop);
                for (int i = 0; i < objects.Length; i++)
                {
                    specListBox.Items.Add(objects[i]);
                }
            }
        }

        private void specListBox_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop) && ((e.AllowedEffect & DragDropEffects.Move) == DragDropEffects.Move))
                e.Effect = DragDropEffects.Move;
        }

        private void clearSpecBtn_Click(object sender, EventArgs e)
        {
            specListBox.Items.Clear();
        }

        private void deleteSpecBtn_Click(object sender, EventArgs e)
        {
            if (specListBox.SelectedIndex >= 0)
                specListBox.Items.RemoveAt(specListBox.SelectedIndex);
        }

        private void formBtn_Click(object sender, EventArgs e)
        {
            if (File.Exists(@"Data\Except.xlsx"))
            {
                Units = new List<Unit>();
                Except = loadDataExcept(@"Data\Except.xlsx");
                for (int i = 0; i < specListBox.Items.Count; i++)
                {
                    Units.AddRange(loadDataSpec(specListBox.Items[i].ToString(), i, unionCheckBox.Checked));
                }
                if (!unionCheckBox.Checked)
                {
                    Units = consolidate(Units);
                    if (groupCheckBox.Checked)
                    {
                        GroupUnits = loadDataGroup(@"Data\база СП.xlsx");
                        Units = findGroup(Units, GroupUnits);
                    }
                    uploadData();
                }
                else uploadData();
            }
            else MessageBox.Show(@"Не найден файл Data\Except.xlsx");
        }

        private List<Unit> consolidate(List<Unit> units)
        {
            for (int i = 0; i < units.Count; i++)
                for (int j = i + 1; j < units.Count; j++)
                    if (units[i].Code != "")
                    {
                        if (units[i].Code == units[j].Code)
                        {
                            if (units[j].Measure.Replace(".", "") == units[i].Measure.Replace(".", ""))
                            {
                                Unit unit = units[i];
                                for (int k = 0; k < unit.cupBoard.Length; k++)
                                {
                                    unit.cupBoard[k].Num += units[j].cupBoard[k].Num;
                                    if (units[j].cupBoard[k].fileName != null || unit.cupBoard[k].fileName == null)
                                        unit.cupBoard[k].fileName = units[j].cupBoard[k].fileName;
                                }
                                units.RemoveAt(j);
                                j--;
                                unit.errorColor = Color.Empty;
                                units[i] = unit;
                            }
                            else
                            {
                                Unit unit = units[j];
                                unit.errorColor = Color.Yellow;
                                units[j] = unit;
                            }
                        }
                        else if (units[i].Name == units[j].Name)
                        {
                            Unit unit = units[j];
                            unit.errorColor = Color.Magenta;
                            units[j] = unit;
                        }

                    }
                    else
                    {
                        if (units[i].Name == units[j].Name)
                        {
                            if (Except.Exists(x => x == units[j].Name))
                            {
                                Unit unit = units[i];
                                for (int k = 0; k < unit.cupBoard.Length; k++)
                                {
                                    unit.cupBoard[k].Num += units[j].cupBoard[k].Num;
                                    if (units[j].cupBoard[k].fileName != null || unit.cupBoard[k].fileName == null)
                                        unit.cupBoard[k].fileName = units[j].cupBoard[k].fileName;
                                }
                                units.RemoveAt(j);
                                j--;
                                unit.errorColor = Color.Empty;
                                units[i] = unit;
                            }
                            else
                            {
                                Unit unit = units[j];
                                unit.errorColor = Color.Red;
                                units[j] = unit;
                            }
                        }
                    }
            return units;
        }

        private List<Unit> findGroup(List<Unit> units, List<GroupUnit> groups)
        {
            for (int i=0; i<units.Count; i++)
            {
                int index = -1;
                if (units[i].Code!="") index = groups.FindIndex(x => x.Code == units[i].Code);
                else index = groups.FindIndex(x => x.Name == units[i].Name);
                if (index >= 0)
                {
                    Unit unit = units[i];
                    unit.Group = groups[index].Group;
                    units[i] = unit;
                }
                else
                {
                    Unit unit = units[i];
                    unit.Group = "";
                    units[i] = unit;
                }
            }
            return units;
        }

        private List<Unit> loadDataSpec(string path, int index, bool union)
        {
            List<Unit> units = new List<Unit>();
            DataSet dataSet = new DataSet("EXCEL");
            string connectionString;
            connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties='Excel 12.0;IMEX=0'";
            OleDbConnection connection = new OleDbConnection(connectionString);
            connection.Open();

            System.Data.DataTable schemaTable = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
            string sheet1 = (string)schemaTable.Rows[0].ItemArray[2];

            string select = String.Format("SELECT * FROM [{0}]", sheet1);
            OleDbDataAdapter adapter = new OleDbDataAdapter(select, connection);
            adapter.Fill(dataSet);
            connection.Close();

            for (int row = 1; row < dataSet.Tables[0].Rows.Count; row++)
            {
                if (dataSet.Tables[0].Rows[row][3].ToString().Length > 0)
                {
                    Unit unit = new Unit();
                    unit.Code = dataSet.Tables[0].Rows[row][2].ToString().Trim();
                    unit.Name = dataSet.Tables[0].Rows[row][3].ToString().Trim();

                    if (union)
                    {
                        unit.Count = double.Parse(dataSet.Tables[0].Rows[row][4].ToString().Trim());
                        unit.FileName = Path.GetFileNameWithoutExtension(specListBox.Items[index].ToString());
                    }
                    else
                    {
                        CupBoard[] cB = new CupBoard[specListBox.Items.Count];
                        for (int i = 0; i < cB.Length; i++)
                        {
                            if (i == index)
                            {
                                cB[index].Num = double.Parse(dataSet.Tables[0].Rows[row][4].ToString().Trim());
                            }
                            else cB[i].Num = 0;
                            cB[i].fileName = Path.GetFileNameWithoutExtension(specListBox.Items[i].ToString());
                        }
                        unit.cupBoard = cB;
                    }
                    unit.Measure = dataSet.Tables[0].Rows[row][5].ToString().Trim();
                    unit.Manufacture = dataSet.Tables[0].Rows[row][6].ToString().Trim();
                    units.Add(unit);
                }
            }
            return units;
        }

        private List<string> loadDataExcept(string path)
        {
            List<string> except = new List<string>();
            DataSet dataSet = new DataSet("EXCEL");
            string connectionString;
            connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties='Excel 12.0;IMEX=0;HDR=NO'";
            OleDbConnection connection = new OleDbConnection(connectionString);
            connection.Open();

            System.Data.DataTable schemaTable = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
            string sheet1 = (string)schemaTable.Rows[0].ItemArray[2];

            string select = String.Format("SELECT * FROM [{0}]", sheet1);
            OleDbDataAdapter adapter = new OleDbDataAdapter(select, connection);
            adapter.Fill(dataSet);
            connection.Close();

            for (int row = 0; row < dataSet.Tables[0].Rows.Count; row++)
                except.Add(dataSet.Tables[0].Rows[row][0].ToString());

            return except;
        }

        private List<GroupUnit> loadDataGroup(string path)
        {
            List<GroupUnit> units = new List<GroupUnit>();
            DataSet dataSet = new DataSet("EXCEL");
            string connectionString;
            connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties='Excel 12.0;IMEX=0;HDR=YES'";
            OleDbConnection connection = new OleDbConnection(connectionString);
            connection.Open();

            System.Data.DataTable schemaTable = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
            string sheet1 = (string)schemaTable.Rows[0].ItemArray[2];

            string select = String.Format("SELECT * FROM [{0}]", sheet1);
            OleDbDataAdapter adapter = new OleDbDataAdapter(select, connection);
            adapter.Fill(dataSet);
            connection.Close();

            for (int row = 0; row < dataSet.Tables[0].Rows.Count; row++)
            {
                GroupUnit unit = new GroupUnit();
                unit.Group = dataSet.Tables[0].Rows[row][0].ToString().Trim();
                unit.Code = dataSet.Tables[0].Rows[row][1].ToString().Trim();
                unit.Name = dataSet.Tables[0].Rows[row][2].ToString().Trim();
                units.Add(unit);
            }
            return units;
        }

        private void uploadData()
        {
            excel = new Excel.Application();
            excel.SheetsInNewWorkbook = 1;
            excel.Workbooks.Add(Type.Missing);
            Excel.Worksheet sheet = (Excel.Worksheet)excel.Sheets.get_Item(1);
            Excel.Range autoFit;

            int curColumn = 1;
            if (groupCheckBox.Checked)
            {
                sheet.Cells[1, curColumn] = "Группа";
                sheet.Columns[curColumn].NumberFormat = "@";
                curColumn++;
            }

            sheet.Cells[1, curColumn] = "Код";
            sheet.Columns[curColumn].NumberFormat = "@";
            curColumn++;

            sheet.Cells[1, curColumn] = "Наименование";
            sheet.Columns[curColumn].NumberFormat = "@";
            curColumn++;

            sheet.Cells[1, curColumn] = "Завод изготовитель";
            sheet.Columns[curColumn].NumberFormat = "@";
            curColumn++;

            sheet.Cells[1, curColumn] = "Ед. изм.";
            sheet.Columns[curColumn].NumberFormat = "@";
            curColumn++;

            if (unionCheckBox.Checked)
            {
                sheet.Cells[1, curColumn] = "Кол-во";
                sheet.Columns[curColumn].NumberFormat = "#";
                curColumn++;
            }

            int curMaxColumn = curColumn - 1;
            for (int i = 0; i < Units.Count; i++)
            {
                if (groupCheckBox.Checked)
                {
                    sheet.Cells[i + 2, curColumn - 5] = Units[i].Group;
                }
                sheet.Cells[i + 2, curColumn - 4] = Units[i].Code;
                sheet.Cells[i + 2, curColumn - 3] = Units[i].Name;
                sheet.Cells[i + 2, curColumn - 2] = Units[i].Manufacture;
                sheet.Cells[i + 2, curColumn - 1] = Units[i].Measure;
                if (!unionCheckBox.Checked)
                {
                    for (int j = 0; j < Units[i].cupBoard.Length; j++)
                    {
                        sheet.Cells[i + 2, curColumn + j] = Units[i].cupBoard[j].Num;
                        if (curColumn + j > curMaxColumn)
                        {
                            curMaxColumn++;
                            sheet.Cells[1, curMaxColumn].NumberFormat = "#";
                            sheet.Cells[1, curMaxColumn] = Units[i].cupBoard[j].fileName;
                        }
                    }
                }
                else
                {
                    sheet.Cells[i + 2, curColumn] = Units[i].Count;
                    sheet.Cells[i + 2, curColumn + 1] = Units[i].FileName;
                    sheet.Cells[1, curColumn + 1] = "Файл";
                    sheet.Columns[curColumn + 1].NumberFormat = "@";
                    curMaxColumn = curColumn + 1;
                    autoFit = (Excel.Range)sheet.Rows[i+2];
                    autoFit.EntireRow.AutoFit();
                    for (int j = 1; j <= curMaxColumn; j++)
                    {
                        autoFit = (Excel.Range)sheet.Columns[j];
                        autoFit.AutoFit();
                    }
                }
            }

            if (!unionCheckBox.Checked)
            {
                sheet.Cells[1, curMaxColumn + 1] = "Сумма по шкафам";
                for (int i = 0; i < Units.Count; i++)
                {
                    Excel.Range c1 = (Excel.Range)sheet.Cells[i + 2, 5];
                    Excel.Range c2 = (Excel.Range)sheet.Cells[i + 2, curMaxColumn];
                    Excel.Range r = (Excel.Range)sheet.Range[c1, c2];
                    ((Excel.Range)sheet.Cells[i + 2, curMaxColumn + 1]).FormulaLocal = "=SUM(" + r.Address.ToString() + ")";
                    autoFit = (Excel.Range)sheet.Cells[i + 2, curMaxColumn + 1];
                    double d = autoFit.Value2;
                    if (d - Math.Truncate(d) != 0)
                    {
                        autoFit = (Excel.Range)sheet.Cells[i + 2, curMaxColumn + 1];
                        autoFit.NumberFormat = "#,#0.0";
                    }
                    autoFit = (Excel.Range)sheet.Rows[i + 2];
                    if (Units[i].errorColor != Color.Empty) autoFit.EntireRow.Interior.Color = Units[i].errorColor;
                    autoFit.EntireRow.AutoFit();
                }
            }
            excel.Visible = true;
        }

        private void unionCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (unionCheckBox.Checked)
            {
                groupCheckBox.Checked = false;
                groupCheckBox.Enabled = false;
            }
            else groupCheckBox.Enabled = true;
        }
    }
}

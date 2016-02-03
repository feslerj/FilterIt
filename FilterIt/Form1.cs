using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel;
using LumenWorks.Framework.IO.Csv;

namespace FilterIt
{
    public partial class frmFilterIt : Form
    {
        private DataTable _csvFileData = null;
        private readonly List<string> _emailEndsWith = new List<string>();
        private readonly List<string> _addressStartsWith = new List<string>();
        private string _fileName = String.Empty;

        public frmFilterIt()
        {
            InitializeComponent();

            _emailEndsWith.AddRange(ConfigurationManager.AppSettings["EmailEndsWith"].Split('|'));
            _addressStartsWith.AddRange(ConfigurationManager.AppSettings["AddressStartsWith"].Split('|'));

            ddFilters.Items.Add("PO Boxes");                    //0
            ddFilters.Items.Add(".edu & .gov Email Addresses"); //1
        }

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var result = openFileDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
                _fileName = openFileDialog1.FileName;

                ddFilters.Visible = btnFilterIt.Visible = false;
                lstColumns.SelectedIndex = -1;
                lstColumns.Items.Clear();
                lblCurrentFile.Text = Path.GetFileName(_fileName);
                
                _csvFileData = _fileName.EndsWith(".csv") ? GetDataSetFromCsvFile(_fileName) : GetDataSetFromExcelFile(_fileName);

                foreach (DataColumn column in _csvFileData.Columns)
                {
                    lstColumns.Items.Add(column.ColumnName);
                }
            }
        }

        private DataTable GetDataSetFromCsvFile(string filename)
        {
            var csvReader = new CachedCsvReader(File.OpenText(filename), true);

            var dt = new DataTable();
            foreach (var column in csvReader.Columns)
            {
                dt.Columns.Add(column.Name);
            }
            var headerRow = dt.NewRow();
            for (int i = 0; i < csvReader.Columns.Count; i++)
            {
                headerRow[i] = csvReader.Columns[i].Name;
            }
            dt.Rows.Add(headerRow);
            while (csvReader.ReadNextRecord())
            {
                var row = dt.NewRow();
                for (int i = 0; i < csvReader.FieldCount; i++)
                {
                    row[i] = csvReader[i];
                }
                dt.Rows.Add(row);
            }
            return dt;
        }

        private DataTable GetDataSetFromExcelFile(string filename)
        {
            FileStream stream = File.Open(filename, FileMode.Open, FileAccess.Read);

            // Reading from a binary Excel file ('97-2003 format; *.xls)
            if (filename.EndsWith(".xls"))
            {
                using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(stream))
                {
                    return GetDataTableFromExcelReader(excelReader);
                }
            }
                // Reading from a OpenXml Excel file (2007 format; *.xlsx)
            else
            {
                using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream))
                {
                    return GetDataTableFromExcelReader(excelReader);
                }
            }
        }

        private DataTable GetDataTableFromExcelReader(IExcelDataReader excelReader)
        {
            var dt = excelReader.AsDataSet().Tables[0];
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                dt.Columns[i].ColumnName = String.IsNullOrEmpty(dt.Rows[0][i].ToString()) ? dt.Columns[i].ColumnName : dt.Rows[0][i].ToString();
            }
            return dt;
        }

        private void lstColumns_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lstColumns.SelectedIndex != -1)
            {
                ddFilters.Visible = btnFilterIt.Visible = true;
            }
        }

        private bool ShouldRemoveRow(int selectedColumnIndex, DataRow row, FilterType filterType)
        {
            string field = row[selectedColumnIndex].ToString().ToLower();

            switch (filterType)
            {
                case FilterType.FilterByAddress:
                    return _addressStartsWith
                        .Select(filterStr => filterStr.ToLower())
                        .Any(field.StartsWith);
                case FilterType.FilterByEmail:
                    return _emailEndsWith
                        .Select(filterStr => filterStr.ToLower())
                        .Any(field.EndsWith);
            }

            return false;
        }

        private void btnFilterIt_Click(object sender, EventArgs e)
        {
            if (_csvFileData == null || ddFilters.SelectedIndex <= -1 || lstColumns.SelectedIndex <= -1)
                return;

            var removedRows = new List<DataRow>();
            var filteredCsvFileData = new DataTable();
            filteredCsvFileData.Columns.AddRange(_csvFileData.Columns.OfType<DataColumn>().Select(col => new DataColumn(col.ColumnName)).ToArray());

            foreach (DataRow row in _csvFileData.Rows)
            {
                if (ShouldRemoveRow(lstColumns.SelectedIndex, row, (FilterType)ddFilters.SelectedIndex))
                {
                    removedRows.Add(row);
                }
                else
                {
                    filteredCsvFileData.ImportRow(row);
                }
            }

            MessageBox.Show(string.Format("Records Found To Be Removed {0}", removedRows.Count));

            SaveRecords(string.Concat(_fileName, "_removed_", DateTime.Now.Ticks, ".csv"), removedRows);

            _csvFileData = filteredCsvFileData;
        }

        private void SaveRecords(string filename, IEnumerable<DataRow> removedRows)
        {
            using (var sw = new StreamWriter(filename))
            {
                foreach (var row in removedRows)
                {
                    sw.WriteLine(WriteCsvLine(row));
                }
                sw.Close();
            }
        }

        private string WriteCsvLine(DataRow row)
        {
            var build = new StringBuilder();
            foreach (var field in row.ItemArray)
            {
                string strField = field.ToString();
                if (strField.Contains(",") || strField.Contains("\""))
                {
                    strField = strField.Replace("\"", "\"\"");
                    strField = string.Concat("\"", strField, "\"");
                }
                build.Append(strField);
                build.Append(",");
            }
            build.Remove(build.Length - 1, 1);
            return build.ToString();
        }

        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string filteredFileName = string.Concat(_fileName, "_filtered_", DateTime.Now.Ticks, ".csv");
            SaveRecords(filteredFileName, _csvFileData.Rows.OfType<DataRow>());

            MessageBox.Show(string.Format("Saved as {0}!", filteredFileName));
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }

    public enum FilterType
    {
        FilterByAddress = 0,
        FilterByEmail = 1,
    }
}

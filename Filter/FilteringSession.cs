using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel;
using LumenWorks.Framework.IO.Csv;

namespace Filter
{
    public class FilteringSession
    {
        private DataTable _fileData = null;
        private DataTable _filteredFileData = null;
        private readonly List<string> _emailEndsWith = new List<string>();
        private readonly List<string> _addressStartsWith = new List<string>();
        private string _fileName = String.Empty;

        public LogState Logger = new LogState();

        public FilteringSession(IEnumerable<string> emailEndsWithFilter, IEnumerable<string> addressStartsWithFilter)
        {
            _emailEndsWith.AddRange(emailEndsWithFilter);
            _addressStartsWith.AddRange(addressStartsWithFilter);
        }

        public string[] GetHeaders()
        {
            if (_fileData == null)
                return new string[0];

            var headers = new string[_fileData.Columns.Count];

            var columnHeaders = GetHeadersAsDataColumns();

            for(var i = 0; i < _fileData.Columns.Count; i++)
            {
                headers[i] = columnHeaders[i].ColumnName;
            }

            return headers;
        }

        public DataColumn[] GetHeadersAsDataColumns()
        {
            return
                _fileData.Columns
                .OfType<DataColumn>()
                .OrderBy(c => c.Ordinal)
                .Select(c => new DataColumn(c.ColumnName))
                .ToArray();
        }

        public bool LoadFile(string fileName)
        {
            _fileName = fileName;

            try
            {
                _fileData = _fileName.EndsWith(".csv") ? GetDataSetFromCsvFile(_fileName) : GetDataSetFromExcelFile(_fileName);
            }
            catch (Exception ex)
            {
                Logger.Error("Could not load file");
                return false;
            }
            return true;
        }

        public bool SaveRecords(string filename, IEnumerable<DataRow> records)
        {
            try
            {
                using (var sw = new StreamWriter(filename))
                {
                    foreach (var row in records)
                    {
                        sw.WriteLine(WriteCsvLine(row));
                    }
                    sw.Close();
                }
            }
            catch (Exception ex)
            {
                Logger.Error("Error saving records", ex);
                return false;
            }

            return true;
        }

        public bool SaveRecords(string filename)
        {
            return SaveRecords(filename, _fileData.Rows.OfType<DataRow>());
        }

        public List<DataRow> Filter(FilterType filterType, int columnIndex)
        {
            var removedRows = new List<DataRow>();

            if (_fileData == null)
            {
                Logger.Error("No filtering done, file to filter has not been loaded.");
                return removedRows;
            }

            if (filterType != FilterType.FilterByAddress && filterType != FilterType.FilterByEmail)
            {
                Logger.Error("No filtering done, invalid filter type.");
                return removedRows;
            }

            if (columnIndex < 0)
            {
                Logger.Error("No filtering done, no column selected.");
                return removedRows;
            }

            var filteredFileData = new DataTable();
            filteredFileData.Columns.AddRange(GetHeadersAsDataColumns());

            foreach (DataRow row in _fileData.Rows)
            {
                if (ShouldRemoveRow(columnIndex, row, filterType))
                {
                    removedRows.Add(row);
                }
                else
                {
                    filteredFileData.ImportRow(row);
                }
            }

            _filteredFileData = filteredFileData;

            return removedRows;
        }

        public int ConfirmFilter()
        {
            if (_filteredFileData == null)
            {
                Logger.Error("No filtering has been done, file data remains unchanged.");
                return 0;
            }

            int count = _fileData.Rows.Count - _filteredFileData.Rows.Count;

            _fileData = _filteredFileData;
            _filteredFileData = null;

            return count;
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
    }
}

﻿using System;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Filter;

namespace FilterIt
{
    public partial class FrmFilterIt : Form
    {
        private string _fileName = String.Empty;
        private FilteringSession _filterSesion = null;

        public FrmFilterIt()
        {
            InitializeComponent();

            //Initial business logic with filters
            _filterSesion = new FilteringSession(
                ConfigurationManager.AppSettings["EmailEndsWith"].Split('|'), 
                ConfigurationManager.AppSettings["AddressStartsWith"].Split('|'));

            //Add items that correspond to FilterItems enum
            ddFilters.Items.Add("PO Boxes and blanks");                 //0
            ddFilters.Items.Add("blank, .edu & .gov Email Addresses");  //1
        }

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var result = openFileDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
                //get name of the file we will be working with
                _fileName = openFileDialog1.FileName;

                //Clear and set all UI information
                lblFilterOptions.Visible = ddFilters.Visible = btnFilterIt.Visible = false;
                lblNumberOfRecords.Text = @"Loading ...";
                lstColumns.SelectedIndex = -1;
                lstColumns.Items.Clear();
                lblCurrentFile.Text = Path.GetFileName(_fileName);

                //actually load up the file into the session
                _filterSesion.LoadFile(_fileName);

                //Load up all of the columns in order in the list box so the user can choose which column to filter on.
                //Note: It looks weird Resharper is yelling about co-varient stuff, 
                //and I'd rather just not hear it instead of having a micro performance increase
                lstColumns.Items.AddRange(_filterSesion.GetHeaders().OfType<object>().ToArray());

                lblNumberOfRecords.Text = _filterSesion.NumberOfRecords.ToString();
            }
        }

        private void lstColumns_SelectedIndexChanged(object sender, EventArgs e)
        {
            //If the selected index is actually picking a column to filter on, make the do-dads visible
            if (lstColumns.SelectedIndex != -1)
            {
                lblFilterOptions.Visible = ddFilters.Visible = btnFilterIt.Visible = true;
            }
        }

        private void btnFilterIt_Click(object sender, EventArgs e)
        {
            //Check to make sure its in a valid state to be filtered
            if (!_filterSesion.FileIsLoaded || ddFilters.SelectedIndex <= -1 || lstColumns.SelectedIndex <= -1)
                return;

            //Capture removed rows for saving seperately
            var removedRows = _filterSesion.Filter((FilterType)ddFilters.SelectedIndex, lstColumns.SelectedIndex);

            string removeRecordsFileName = string.Concat(_fileName, "_removed_", DateTime.Now.Ticks, ".csv");
            _filterSesion.SaveRecords(removeRecordsFileName, removedRows);

            var result = MessageBox.Show(string.Format("Records Found To Be Removed {0}.{1}Inspect removed records here: {2}", 
                                                        removedRows.Count, Environment.NewLine, removeRecordsFileName), 
                                        @"Remove Records?", MessageBoxButtons.YesNo);

            //Only save state if filtering is confirmed
            if (result == DialogResult.Yes)
            {
                _filterSesion.ConfirmFilter();
                lblNumberOfRecords.Text = _filterSesion.NumberOfRecords.ToString();
            }
        }

        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string filteredFileName = string.Concat(_fileName, "_filtered_", DateTime.Now.Ticks, ".csv");
            _filterSesion.SaveRecords(filteredFileName);
            
            MessageBox.Show(string.Format("Saved as {0}!", filteredFileName));

            if (Environment.OSVersion.Platform == PlatformID.Win32Windows || Environment.OSVersion.Platform == PlatformID.Win32NT)
            {
                System.Diagnostics.Process.Start("explorer.exe", @"/select, " + filteredFileName);
            }
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void helpToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show(String.Concat(@"1. Open a csv or Excel file you would like to filter records on.", Environment.NewLine,
                                          @"2. Choose a column to filter in the list box", Environment.NewLine,
                                          @"3. Choose a way to filter the item in the dropdown box", Environment.NewLine,
                                          @"4. Push the Filter It! button to save a list of the removed records and confirm that you want those records removed.", Environment.NewLine,
                                          @"5. Save the new filtered set of records along side the original file"
                                ), "Info", MessageBoxButtons.OK);
            ;
        }
    }
}

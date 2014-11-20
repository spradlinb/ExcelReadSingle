using Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelReadSingle
{
    public partial class ExcelProcessor : Form
    {
        private Row[] _rows = null;
        private int _rowMax = 0;
        private int _colMax = 0;
        private int _rowCount = 0;
        private int _colCount = 0;

        public ExcelProcessor()
        {
            InitializeComponent();
        }

        private void browseButton_Click(object sender, EventArgs e)
        {
            if (_rowMax != 0 && _colMax != 0)
            {
                _rowMax = 0;
                _colMax = 0;
                _rowCount = 0;
                _colCount = 0;
                nextButton.Enabled = true;
            }

            openFileDialog1.Filter = "Excel Files (*.xls,*.xlsx)|*.xls;*.xlsx";
            DialogResult result = openFileDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
                string fileOpened = openFileDialog1.FileName;
                try
                {
                    filePath.Text = fileOpened;
                    filePath.Refresh();
                    var worksheets = Workbook.Worksheets(fileOpened);
                    var firstWorksheet = worksheets.FirstOrDefault();
                    _rows = firstWorksheet.Rows;

                    if (_rows != null && _rows.Length > 1)
                    {
                        _rowMax = _rows.Length - 1;
                        _colMax = _rows[0].Cells.Length - 1;
                        nextButton_Click(this, e);
                    }
                }
                catch (Exception ex)
                {

                }
            }
        }

        private void nextButton_Click(object sender, EventArgs e)
        {
            if (_rowCount >= _rowMax && _colCount >= _colMax)
            {
                nextButton.Enabled = false;
            }

            var thisRow = _rows[_rowCount];
            var thisCol = thisRow.Cells[_colCount];

            cellValue.Text = thisCol.Text;
            cellValue.SelectAll();
            cellValue.Focus();

            if (_colCount < _colMax)
            {
                _colCount++;
            }
            else
            {
                _colCount = 0;
                _rowCount++;
            }
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }
    }
}

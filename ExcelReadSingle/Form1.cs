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
        private enum ButtonsList { Next, Previous }
        private ButtonsList? _lastClicked = null;

        public ExcelProcessor()
        {
            InitializeComponent();
            previousButton.Enabled = false;
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
            rowLabel.Text = string.Format("Row: {0}", _rowCount + 1);
            if (previousButton.Enabled == false && _colCount > 0)
                previousButton.Enabled = true;

            if (_lastClicked == ButtonsList.Previous && _colCount == _colMax)
            {
                _rowCount++;
                _colCount = 0;
                rowLabel.Text = string.Format("Row: {0}", _rowCount + 1);
            }

            else if (_lastClicked == ButtonsList.Previous && _colCount < _colMax)
                _colCount++;
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
            _lastClicked = ButtonsList.Next;
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void previousButton_Click(object sender, EventArgs e)
        {
            if (_lastClicked == ButtonsList.Next)
                _colCount--;
            if (_colCount <= 0)
            {
                _rowCount--;
                _colCount = _colMax;
            }
            else
            {
                _colCount--;
            }

            var thisRow = _rows[_rowCount];
            var thisCol = thisRow.Cells[_colCount];

            cellValue.Text = thisCol.Text;
            cellValue.SelectAll();
            cellValue.Focus();
            _lastClicked = ButtonsList.Previous;

            if (_rowCount == 0 && _colCount == 0)
            {
                previousButton.Enabled = false;
            }

            rowLabel.Text = string.Format("Row: {0}", _rowCount+1);

            if (nextButton.Enabled == false && _colCount < _colMax)
                nextButton.Enabled = true;
        }
    }
}

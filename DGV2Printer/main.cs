using System;
using System.Windows.Forms;
using System.Drawing;
using System.Drawing.Printing;
using System.Data;
using System.Collections;


namespace DGV2Printer
{
    internal class DataGridPrinter
    {

        private PrintDocument ThePrintDocument;
        //private DataTable TheTable;
        private DataGridView TheDataGrid;

        public bool isRTL { get; set; }

        public string RHeader { get; set; }
        public string RFooter { get; set; }

        public int RowCount = 0;  // current count of rows;
        private const int kVerticalCellLeeway = 15;
        public int PageNumber = 1;
        public ArrayList Lines = new ArrayList();

        int PageWidth;
        int PageHeight;
        int TopMargin;
        int BottomMargin;
        int LeftMargin;

        int columnsCount = 0;

        int headerHeight = 40;


        public DataGridPrinter(DataGridView aGrid, PrintDocument aPrintDocument)
        {
            TheDataGrid = aGrid;

            //remove unvisible columns from GridView
            for (int i = 0; i < TheDataGrid.Columns.Count; i++)
            {
                if (TheDataGrid.Columns[i].Visible == false)
                {
                    TheDataGrid.Columns.RemoveAt(i);
                    i--;
                }
            }

            columnsCount = TheDataGrid.Columns.Count;


            ThePrintDocument = aPrintDocument;

            PageWidth = (int)ThePrintDocument.DefaultPageSettings.PrintableArea.Width;
            PageHeight = (int)ThePrintDocument.DefaultPageSettings.PrintableArea.Height;

            TopMargin = (int)ThePrintDocument.DefaultPageSettings.PrintableArea.Y;
            BottomMargin = ThePrintDocument.DefaultPageSettings.PaperSize.Height - PageHeight - TopMargin;
            LeftMargin = 0;// (int)(ThePrintDocument.DefaultPageSettings.PrintableArea.X / 2);
        }

        public void DrawReportHeader(Graphics g)
        {
            if (RHeader != null)
            {
                g.FillRectangle(new SolidBrush(Color.LightGray), 50, TopMargin, PageWidth - 100, TopMargin + headerHeight - 10);
                g.DrawRectangle(new Pen(Color.Black), 50, TopMargin, PageWidth - 100, TopMargin + headerHeight - 10);
                StringFormat sformat = new StringFormat();
                sformat.LineAlignment = StringAlignment.Center;
                sformat.Alignment = StringAlignment.Center;
                g.DrawString(RHeader, TheDataGrid.ColumnHeadersDefaultCellStyle.Font, new SolidBrush(Color.Black), new RectangleF(LeftMargin, TopMargin, PageWidth, 30), sformat);
            }
            else
                headerHeight = 10;
        }

        public void DrawHeader(Graphics g)
        {
            DrawReportHeader(g);
            SolidBrush ForeBrush = new SolidBrush(TheDataGrid.ColumnHeadersDefaultCellStyle.ForeColor);
            SolidBrush BackBrush = new SolidBrush(TheDataGrid.ColumnHeadersDefaultCellStyle.BackColor);
            Pen TheLinePen = new Pen(Color.Black, 1);
            StringFormat cellformat = new StringFormat();
            cellformat.Trimming = StringTrimming.EllipsisCharacter;
            cellformat.FormatFlags = StringFormatFlags.NoWrap | StringFormatFlags.LineLimit;
            cellformat.LineAlignment = StringAlignment.Center;

            cellformat.Alignment = (isRTL) ? StringAlignment.Far : StringAlignment.Near;


            int columnwidth = (PageWidth) / columnsCount;

            int initialRowCount = RowCount;

            g.DrawLine(TheLinePen, LeftMargin, headerHeight + TopMargin, LeftMargin + PageWidth, headerHeight + TopMargin);

            // draw the table header
            float startxposition = LeftMargin;
            RectangleF nextcellbounds = new RectangleF(0, 0, 0, 0);

            RectangleF HeaderBounds = new RectangleF(0, 0, 0, 0);

            HeaderBounds.X = LeftMargin;
            HeaderBounds.Y = headerHeight + TopMargin + (RowCount - initialRowCount) * (TheDataGrid.Font.SizeInPoints + kVerticalCellLeeway);
            HeaderBounds.Height = TheDataGrid.Font.SizeInPoints + kVerticalCellLeeway;
            HeaderBounds.Width = PageWidth;

            g.FillRectangle(BackBrush, HeaderBounds);

            float sumOfWeight = 0;
            for (int jj = 0; jj < columnsCount; jj++) sumOfWeight += TheDataGrid.Columns[jj].FillWeight;
            float weightPerCell = (PageWidth / sumOfWeight);

            for (int k = 0; k < columnsCount; k++)
            {
                int idex = (isRTL) ? columnsCount - 1 - k : k;
                string nextcolumn = TheDataGrid.Columns[idex].HeaderText.ToString();

                float thisComunWidth = weightPerCell * TheDataGrid.Columns[idex].FillWeight;

                RectangleF cellbounds = new RectangleF(startxposition, headerHeight + TopMargin + (RowCount - initialRowCount) * (TheDataGrid.Font.SizeInPoints + kVerticalCellLeeway),
                    thisComunWidth,
                    TheDataGrid.ColumnHeadersDefaultCellStyle.Font.SizeInPoints + kVerticalCellLeeway);
                nextcellbounds = cellbounds;

                //if (startxposition + thisComunWidth <= LeftMargin + PageWidth)
                {
                    g.DrawString(nextcolumn, TheDataGrid.ColumnHeadersDefaultCellStyle.Font, ForeBrush, cellbounds, cellformat);
                }

                startxposition = startxposition + thisComunWidth;

            }

            //if (TheDataGrid != DataGridLineStyle.None)
            g.DrawLine(TheLinePen, LeftMargin, nextcellbounds.Bottom, LeftMargin + PageWidth, nextcellbounds.Bottom);
        }

        public bool DrawRows(Graphics g)
        {
            int lastRowBottom = TopMargin;

            //  try
            {
                Lines.Clear();
                SolidBrush ForeBrush = new SolidBrush(TheDataGrid.ForeColor);
                SolidBrush BackBrush = new SolidBrush(TheDataGrid.BackColor);
                SolidBrush AlternatingBackBrush = new SolidBrush(Color.White);
                Pen TheLinePen = new Pen(Color.Black, 1);
                StringFormat cellformat = new StringFormat();
                cellformat.Trimming = StringTrimming.EllipsisCharacter;
                cellformat.FormatFlags = StringFormatFlags.NoWrap | StringFormatFlags.LineLimit;
                cellformat.LineAlignment = StringAlignment.Center;
                cellformat.Alignment = StringAlignment.Center;


                int initialRowCount = RowCount;

                RectangleF RowBounds = new RectangleF(0, 0, 0, 0);

                float sumOfWeight = 0;
                for (int jj = 0; jj < columnsCount; jj++) sumOfWeight += TheDataGrid.Columns[jj].FillWeight;
                float weightPerCell = (PageWidth / sumOfWeight);

                // draw vertical lines

                // draw the rows of the table
                for (int i = initialRowCount; i < TheDataGrid.Rows.Count; i++)
                {
                    DataGridViewRow dr = TheDataGrid.Rows[i];
                    int startxposition = LeftMargin;

                    RowBounds.X = LeftMargin;
                    RowBounds.Y = headerHeight + TopMargin + ((RowCount - initialRowCount) + 1) * (TheDataGrid.Font.SizeInPoints + kVerticalCellLeeway);
                    RowBounds.Height = TheDataGrid.Font.SizeInPoints + kVerticalCellLeeway;
                    RowBounds.Width = PageWidth;
                    Lines.Add(RowBounds.Bottom);

                    if (i % 2 == 0)
                    {
                        g.FillRectangle(BackBrush, RowBounds);
                    }
                    else
                    {
                        g.FillRectangle(new SolidBrush(Color.LightGray), RowBounds);
                    }


                    for (int j = 0; j < columnsCount; j++)
                    {
                        int idex = (isRTL) ? columnsCount - 1 - j : j;

                        float columnwidth = weightPerCell * TheDataGrid.Columns[idex].FillWeight;

                        RectangleF cellbounds = new RectangleF(startxposition,
                            headerHeight + TopMargin + ((RowCount - initialRowCount) + 1) * (TheDataGrid.Font.SizeInPoints + kVerticalCellLeeway),
                            columnwidth,
                            TheDataGrid.Font.SizeInPoints + kVerticalCellLeeway);

                        {
                            string st = " ";
                            if (dr.Cells[idex].Value != null)
                            {
                                try
                                {
                                    if (TheDataGrid.Columns[idex].DefaultCellStyle.Format == "N0")
                                        st = Convert.ToInt64(dr.Cells[idex].Value.ToString()).ToString("#,##");
                                    else
                                        st = dr.Cells[idex].Value.ToString();
                                }
                                catch { st = dr.Cells[idex].Value.ToString(); }
                            }
                            if (isRTL)
                            {
                                if (st == "True") st = "بله";
                                if (st == "False") st = "خیر";
                            }
                            g.DrawString(st, TheDataGrid.Font, ForeBrush, cellbounds, cellformat);
                            lastRowBottom = (int)cellbounds.Bottom;
                        }

                        startxposition = (int)(startxposition + columnwidth);
                    }

                    RowCount++;

                    //if (RowCount * (TheDataGrid.Font.SizeInPoints + kVerticalCellLeeway) > (PageHeight * PageNumber) - (BottomMargin + TopMargin))
                    if (RowCount * (TheDataGrid.Font.SizeInPoints + kVerticalCellLeeway) > ((PageHeight * PageNumber) - BottomMargin - headerHeight - 50))
                    {
                        DrawHorizontalLines(g, Lines);
                        DrawVerticalGridLines(g, TheLinePen, lastRowBottom);
                        DrawReportFooter(g, lastRowBottom);
                        return true;
                    }
                }

                DrawHorizontalLines(g, Lines);
                DrawVerticalGridLines(g, TheLinePen, lastRowBottom);
                DrawReportFooter(g, lastRowBottom);

                return false;

            }
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message.ToString());
            //    return false;
            //}

        }


        public void DrawReportFooter(Graphics g, int bottomY)
        {
            if (RFooter != null)
            {
                g.FillRectangle(new SolidBrush(Color.LightGray), 30, bottomY + 10, PageWidth - 60, 30);
                g.DrawRectangle(new Pen(Color.Black), 30, bottomY + 10, PageWidth - 60, 30);
                StringFormat sformat = new StringFormat();
                sformat.LineAlignment = StringAlignment.Center;
                sformat.Alignment = (isRTL) ? StringAlignment.Far : StringAlignment.Near;
                g.DrawString(RFooter, TheDataGrid.ColumnHeadersDefaultCellStyle.Font, new SolidBrush(Color.Black), new RectangleF(50, bottomY + 10, PageWidth - 100, 30), sformat);
            }
        }

        void DrawHorizontalLines(Graphics g, ArrayList lines)
        {
            Pen TheLinePen = new Pen(Color.Black, 1);

            for (int i = 0; i < lines.Count; i++)
                g.DrawLine(TheLinePen, LeftMargin, (float)lines[i], LeftMargin + PageWidth, (float)lines[i]);
        }

        void DrawVerticalGridLines(Graphics g, Pen TheLinePen, int bottom)
        {
            float sumOfWeight = 0;
            for (int jj = 0; jj < columnsCount; jj++) sumOfWeight += TheDataGrid.Columns[jj].FillWeight;
            float weightPerCell = PageWidth / sumOfWeight;
            float ss = 0;
            for (int k = 0; k <= columnsCount; k++)
            {
                int idex = (isRTL) ? columnsCount - 1 - k : k;
                float columnwidth = 0;
                try { columnwidth = weightPerCell * TheDataGrid.Columns[idex].FillWeight; }
                catch { columnwidth = PageWidth; }
                g.DrawLine(TheLinePen, LeftMargin + ss, headerHeight + TopMargin, LeftMargin + ss, bottom);
                ss += columnwidth;
            }
        }


        public bool DrawDataGrid(Graphics g)
        {
            // try
            {
                DrawHeader(g);
                bool bContinue = DrawRows(g);
                return bContinue;
            }
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message.ToString());
            //    return false;
            //}
        }
    }

    /// <summary>
    /// TinarSoft DataGridView Printer  http://www.tinarsoft.com
    /// </summary>
    public class PrintDataGridView
    {
        /// <summary>
        /// Is This Report RTL (like Persian,Arabic,..... language)
        /// </summary>
        [System.ComponentModel.DefaultValue(false)]
        public bool isRightToLeft { get; set; }

        /// <summary>
        /// Report Header , Printed in Top Of Report Pages
        /// </summary>
        public string ReportHeader { get; set; }

        /// <summary>
        /// Report Footer , Printed in Bottom Of Report Pages
        /// </summary>
        public string ReportFooter { get; set; }


        private DataGridView myDGV;
        private System.Drawing.Printing.PrintDocument printDocument1;
        private System.Windows.Forms.PrintDialog printDialog1;

        public PrintDataGridView(DataGridView myDataGridView)
        {
            this.myDGV = myDataGridView;
            this.printDocument1 = new System.Drawing.Printing.PrintDocument();
            this.printDialog1 = new System.Windows.Forms.PrintDialog();

            this.printDocument1.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(this.printDocument1_PrintPage);
            this.printDialog1.AllowSelection = true;
            this.printDialog1.AllowSomePages = true;
            this.printDialog1.Document = this.printDocument1;
        }

        private DataGridPrinter dataGridPrinter1 = null;
        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            Graphics g = e.Graphics;
            //DrawTopLabel(g);
            bool more = dataGridPrinter1.DrawDataGrid(g);
            if (more == true)
            {
                e.HasMorePages = true;
                dataGridPrinter1.PageNumber++;
            }
        }


        /// <summary>
        /// Start Printing The Report
        /// </summary>
        public void Print()
        {
            if (printDialog1.ShowDialog() == DialogResult.OK)
            {
                dataGridPrinter1 = new DataGridPrinter(myDGV, printDocument1);
                dataGridPrinter1.RHeader = ReportHeader;
                dataGridPrinter1.isRTL = isRightToLeft;
                dataGridPrinter1.RFooter = ReportFooter;
                printDocument1.Print();
            }
        }
    }
}


using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;
using Word = Microsoft.Office.Interop.Word;

namespace MartinAppGUI
{
    public partial class DataHandlerApp : Form
    {
        List<Row> originalData;
        List<Row> modifiedData;
        List<Row> orderedData;
        int sumOfModifiedData;
        int sumOfOriginalData;
        string path = string.Empty;
        string templateFilePath = string.Empty;
        int completedWordDoc = 0;
        private const int buttondown = 161;
        private const int HT_CAPTION = 2;

        public DataHandlerApp()
        {
            InitializeComponent();
            createTableColumns();
            
            modifiedData = new List<Row>();
            orderedData = new List<Row>();
            bttnStop.Enabled = false;
            bttnExport.Enabled = false;       
            bttnOpen.FlatAppearance.BorderSize = 0;
            bttnExport.FlatAppearance.BorderSize = 0;   
            bttnClose.FlatAppearance.BorderSize = 0;
            bttnStop.FlatAppearance.BorderSize = 0;
            button1.FlatAppearance.BorderSize = 0;
        }

        //Button events
        private void megnyitas_Click(object sender, EventArgs e)
        {
            var filePath = string.Empty;

            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = System.Windows.Forms.Application.StartupPath;
                openFileDialog.Filter = "xlsx files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    filePath = openFileDialog.FileName;
                    readFromExcel(filePath);
                    if (backgroundWorker1.IsBusy != true)
                        backgroundWorker1.RunWorkerAsync(sender);
                }
            }
        }
        private void exportbutton_Click(object sender, EventArgs e)
        {
            folderBrowser.Description = "Mentés helye:";
            folderBrowser.SelectedPath = System.Windows.Forms.Application.StartupPath;

            if (folderBrowser.ShowDialog() == DialogResult.OK)
            {
                if (folderBrowser.SelectedPath != null)
                {
                    path = folderBrowser.SelectedPath;

                    if (backgroundWorker2.IsBusy != true)
                    {
                        bttnStop.Enabled = true;
                        bttnExport.Enabled = false;

                        backgroundWorker2.RunWorkerAsync(sender);
                    }
                }
            }


        }
        private void bttnDelete_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = null;
            dataGridView1.Rows.Clear();
            bttnOpen.Enabled = true;
          
            bttnExport.Enabled = false;
            
        }
        private void button1_Click(object sender, EventArgs e)
        {
            backgroundWorker2.CancelAsync();
            bttnStop.Enabled = false;
            bttnExport.Enabled = true;
        }

        //BackgrondWorker events
        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            backgroundWorker1.ReportProgress(0, "Working...");

            modifiedData.Add(originalData[0]);

            for (int i = 1; i < sumOfOriginalData; ++i)
            {

                if (!isAlreadyAdded(originalData[i], modifiedData))
                {
                    modifiedData.Add(originalData[i]);
                }
                else
                {
                    addToExsistingRecord(originalData[i], modifiedData);
                }

            }
            sumOfModifiedData = modifiedData.Count();

            for (int i = 0; i < sumOfModifiedData; ++i)
            {
                StringBuilder sb = new StringBuilder();

                for (int j = 0; j < modifiedData[i].oktatok.Count(); ++j)
                {
                    if (j < modifiedData[i].oktatok.Count()-1)
                        sb.Append(modifiedData[i].oktatok[j] + ", ");
                    else
                        sb.Append(modifiedData[i].oktatok[j]);
                }

                backgroundWorker1.ReportProgress((100 * i) / sumOfModifiedData, "Working...");

                string osszOktato = sb.ToString();

                string[] row = { (i + 1).ToString(), modifiedData[i].targyKod, modifiedData[i].targyNev, modifiedData[i].letszam.ToString(), osszOktato };
                dataGridView1.Invoke(new Action(() => dataGridView1.Rows.Add(row)));

            }
            backgroundWorker1.ReportProgress(100, "Complete!");
        }
        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar1.Value = e.ProgressPercentage;

        }
        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error != null)
            {
                MessageBox.Show(e.Error.ToString(), "Hiba!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                progressBar1.Value = 0;
            }
            else
            {
                MessageBox.Show("Az adatok feldolgozását befejeztük!", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                bttnExport.Enabled = true;
                progressBar1.Value = 0;
                bttnOpen.Enabled = false;
              
                bttnOpen.Enabled = false;
                originalData = null;      
            }
        }

        private void backgroundWorker2_DoWork(object sender, DoWorkEventArgs e)
        {


            //Parallel.ForEach(modifiedData, (currentRow) =>
            //{
            //    if (backgroundWorker2.CancellationPending == true)
            //    {
            //        e.Cancel = true;
            //    }
            //    else
            //    {

            //        string modifiedTargynev = removeSpecialCharacters(currentRow.targyNev);

            //        if (modifiedTargynev.Length > 100)
            //        {
            //            modifiedTargynev = modifiedTargynev.Remove(100) + "...";
            //        }

            //        string saveTo = path + "\\" + modifiedTargynev + " " + currentRow.targyKod + ".docx";


            //        CreateWordDocument(saveTo, currentRow.targyNev, currentRow.targyKod, (currentRow.letszam).ToString(), currentRow.oktatok);

            //        backgroundWorker2.ReportProgress(1);
            //    }
            //});


            for (int i = 0; i < sumOfModifiedData; ++i)
            {
                if ((backgroundWorker2.CancellationPending == true))
                {
                    e.Cancel = true;
                }
                else
                {

                    int progress = (int)(((float)(i + 1) / sumOfModifiedData) * 100);

                    string modifiedTargynev = removeSpecialCharacters(modifiedData[i].targyNev);

                    if (modifiedTargynev.Length > 100)
                    {
                        modifiedTargynev = modifiedTargynev.Remove(100) + "...";
                    }

                    string saveTo = path + "\\" + modifiedTargynev + " " + modifiedData[i].targyKod + ".docx";


                    CreateWordDocument(saveTo, modifiedData[i].targyNev, modifiedData[i].targyKod, (modifiedData[i].letszam).ToString(), modifiedData[i].oktatok);

                   

                    backgroundWorker2.ReportProgress(1);
                }
            }
        }
        private void backgroundWorker2_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            completedWordDoc += e.ProgressPercentage;

            int progress = (int)(((float)completedWordDoc / sumOfModifiedData) * 100);

            progressBar1.Value = progress;
        }
        private void backgroundWorker2_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {

            if (e.Error != null)
            {
                MessageBox.Show(e.Error.ToString(), "Hiba!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else if (e.Cancelled)
            {
                MessageBox.Show("Sikeres megszakítás!", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                progressBar1.Value = 0;
            }
            else
            {
                MessageBox.Show("Az összes dokumentumot feldolgoztuk!", "Kész", MessageBoxButtons.OK, MessageBoxIcon.Information);
                bttnStop.Enabled = false;
                bttnExport.Enabled = true;
            }

        }

        private void createTableColumns()
        {
            dataGridView1.AutoGenerateColumns = false;
            dataGridView1.AutoGenerateColumns = false;
            dataGridView1.RowHeadersVisible = false;
            dataGridView1.MultiSelect = false;
            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            dataGridView1.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dataGridView1.Columns.Add(new DataGridViewTextBoxColumn()
            {
                HeaderText = "Id",
                ReadOnly = true,
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
                FillWeight = 25
            });
            dataGridView1.Columns.Add(new DataGridViewTextBoxColumn()
            {
                HeaderText = "Tárgykód",
                ReadOnly = true,
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
                FillWeight = 30
            });
            dataGridView1.Columns.Add(new DataGridViewTextBoxColumn()
            {
                HeaderText = "Tárgynév",
                ReadOnly = true,
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
                FillWeight = 75
            });
            dataGridView1.Columns.Add(new DataGridViewTextBoxColumn()
            {
                HeaderText = "Létszám",
                ReadOnly = true,
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
                FillWeight = 30
            });
            dataGridView1.Columns.Add(new DataGridViewTextBoxColumn()
            {
                HeaderText = "Oktatók",
                ReadOnly = true,
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
                FillWeight = 165,

            });
        }
        private void readFromExcel(string path)
        {
            Row row;

            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            using (ExcelPackage xlPackage = new ExcelPackage(new FileInfo(path)))
            {
                var myWorksheet = xlPackage.Workbook.Worksheets.First();
                var totalRows = myWorksheet.Dimension.End.Row;
                var totalColumns = myWorksheet.Dimension.End.Column;
                originalData = new List<Row>();

                for (int rowNum = 2; rowNum <= totalRows; rowNum++)
                {
                    var current_row = myWorksheet.Cells[rowNum, 1, rowNum, totalColumns].Select(c => c.Value == null ? string.Empty : c.Value.ToString()).ToList();
                    row = new Row();
                    ++sumOfOriginalData;
                    row.id = sumOfOriginalData;
                    row.targyKod = current_row[0];
                    row.targyNev = current_row[1];
                    row.letszam = Convert.ToInt32(current_row[2]);

                    foreach (var item in splitOktatokString(current_row[3]))
                    {
                        row.oktatok.Add(item);
                    }
                    originalData.Add(row);
                }
            }
        }
        private List<string> splitOktatokString(string oktatok)
        {
            List<string> tempOktatok = oktatok.Split(new string[] { ", " }, StringSplitOptions.None).ToList();
            return tempOktatok;
        }
        public string removeSpecialCharacters(string str)
        {
            StringBuilder sb = new StringBuilder();
            foreach (char c in str)
            {
                if (c != ':' && c != '/' && c != '?' && c != '<' && c != '>')
                {
                    sb.Append(c);
                }
            }
            return sb.ToString();
        }
        private static bool isAlreadyAdded(Row row, List<Row> modifiedData)
        {
            foreach (var item in modifiedData)
            {
                if (item.targyKod == row.targyKod)
                {
                    return true;
                }
            }
            return false;
        }
        private static void addToExsistingRecord(Row row, List<Row> modifiedData)
        {
            foreach (var item in modifiedData)
            {
                if (item.targyKod == row.targyKod)
                {


                    foreach (var oktato in row.oktatok)
                    {
                        if (!item.oktatok.Contains(oktato))
                        {
                            item.oktatok.Add(oktato);
                        }
                    }

                    if (!item.targyNev.Contains(row.targyNev))
                    {
                        item.targyNev += $", {row.targyNev}";
                    }

                    item.letszam += row.letszam;

                }
            }
        }
        private void CreateWordDocument(object SaveAs, string tantargy, string targykod, string letszam, List<string> oktatok)
        {

            object oMissing = Missing.Value;
            object oEndOfDoc = "\\endofdoc"; /* \endofdoc egy előredefiniált könyvjelző (a doksi végét jelöli) */

            //Word indítása, és egy új Word dokumentum létrehozása
            Word.Application oWord;
            Word.Document oDoc;
            oWord = new Word.Application();
            oWord.Visible = false;
            oDoc = oWord.Documents.Add(ref oMissing, ref oMissing, ref oMissing, ref oMissing);

            //Itt kezdődik a doksiba írás

            //2x2-es táblázat beillesztése, és adatokkal való felttöltése, formázása,

            Word.Table oTable;
            Word.Range wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oTable = oDoc.Tables.Add(wrdRng, 2, 2, ref oMissing, ref oMissing);
            oTable.Range.ParagraphFormat.SpaceAfter = 6;
            oTable.Range.ParagraphFormat.SpaceBefore = 6;
            oTable.Range.Font.Size = 12;
            oTable.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            oTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            oTable.Range.Font.Name = "Times New Roman";
            oTable.Range.Font.Bold = 1;
            oTable.Cell(1, 2).Range.ParagraphFormat.Alignment =
                    oTable.Cell(2, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

            int r, c;

            for (r = 1; r <= 2; r++)
                for (c = 1; c <= 2; c++)
                    if (r == 1 && c == 1)
                        oTable.Cell(r, c).Range.Text = "Tantárgy";
                    else if (r == 1 && c == 2)
                        oTable.Cell(r, c).Range.Text = tantargy;
                    else if (r == 2 && c == 1)
                        oTable.Cell(r, c).Range.Text = "Tárgykód";
                    else if (r == 2 && c == 2)
                        oTable.Cell(r, c).Range.Text = targykod;

            //Tárgyat oktatók bekezdés

            Word.Paragraph oPara3;
            object oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oPara3 = oDoc.Content.Paragraphs.Add(ref oRng);
            oPara3.Range.Text = "A tárgyat oktatók:";
            oPara3.Range.Font.Size = 14;
            oPara3.Range.Font.Name = "Times New Roman";
            oPara3.Range.Font.Bold = 1;
            oPara3.Format.SpaceBefore = 24;
            oPara3.Format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            oPara3.Range.InsertParagraphAfter();

            //Oktatók listája

            Word.Paragraph oPara4;
            oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oPara4 = oDoc.Content.Paragraphs.Add(ref oRng);
            oPara4.Range.Font.Size = 12;
            oPara4.Format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            oPara4.Range.ListFormat.ApplyBulletDefault();


            for (int i = 0; i < oktatok.Count(); i++) 
            { 
                if (i < oktatok.Count() - 1)
                    oktatok[i] = oktatok[i] + "\n";
                oPara4.Range.InsertBefore(oktatok[i]);
            }

            //Üres sor (A táblázat előtti távolság)

            //Word.Paragraph oPara5;
            //oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            //oPara5 = oDoc.Content.Paragraphs.Add(ref oRng);
            //oPara5.Range.Text = "";
            //oPara5.Format.SpaceBefore = 32;
            //oPara5.Range.InsertParagraphAfter();

            //Véleményezési adatok táblázat

            wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oTable = oDoc.Tables.Add(wrdRng, 2, 4, ref oMissing, ref oMissing);
            oTable.Range.ParagraphFormat.SpaceBefore = 6;
            oTable.Range.Font.Size = 12;
            oTable.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            oTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            oTable.Range.Font.Name = "Times New Roman";
            oTable.Rows[1].Range.Font.Bold = 1;
            oTable.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;


            for (r = 1; r <= 2; r++)
                for (c = 1; c <= 4; c++)
                    if (r == 1 && c == 1)
                        oTable.Cell(r, c).Range.Text = "Véleményezésre felkértek száma (fő)";
                    else if (r == 1 && c == 2)
                        oTable.Cell(r, c).Range.Text = "Véleményező jogosultak (fő)";
                    else if (r == 1 && c == 3)
                        oTable.Cell(r, c).Range.Text = "Véleményezők (fő)";
                    else if (r == 1 && c == 4)
                        oTable.Cell(r, c).Range.Text = "Véleményezői arány a jogosultakhoz képest (%)";
                    else if (r == 2 && c == 1)
                        oTable.Cell(r, c).Range.Text = letszam;


            //1. ábra. A tanórákon való részvételi arány

            Word.Paragraph oPara6;
            oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oPara6 = oDoc.Content.Paragraphs.Add(ref oRng);
            oPara6.Range.Text = "1. ábra. A tanórákon való részvételi arány";
            oPara6.Range.Font.Size = 14;
            oPara6.Range.Font.Bold = 1;
            oPara6.Format.SpaceBefore = 120;
            oPara6.Format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            oPara6.Range.InsertParagraphAfter();

            //2. ábra. Tantárgyi elégedettség (százalékos elosztás)

            Word.Paragraph oPara7;
            oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oPara7 = oDoc.Content.Paragraphs.Add(ref oRng);
            oPara7.Range.Text = "2. ábra. Tantárgyi elégedettség (százalékos elosztás)";
            oPara6.Format.SpaceBefore = 132;
            oPara7.Range.InsertParagraphAfter();


            //Mentés
            oDoc.SaveAs2(ref SaveAs, ref oMissing, ref oMissing, ref oMissing,
                            ref oMissing, ref oMissing, ref oMissing,
                            ref oMissing, ref oMissing, ref oMissing,
                            ref oMissing, ref oMissing, ref oMissing,
                            ref oMissing, ref oMissing, ref oMissing);

            //Doksi bezárása, kilépés
            oDoc.Close();
         

            if (Process.GetProcessesByName("winword").Any())
            {
                oWord = (Word.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Word.Application");
                oWord.Quit();
            }

        }


        private void bttnStop_MouseEnter(object sender, EventArgs e)
        {
            bttnStop.BackColor = Color.FromArgb(30, 84, 161);
        }

        private void bttnExport_MouseEnter(object sender, EventArgs e)
        {
            bttnExport.BackColor = Color.FromArgb(30, 84, 161);
        }


        private void bttnOpen_MouseEnter_1(object sender, EventArgs e)
        {
            bttnOpen.BackColor = Color.FromArgb(30, 84, 161);
        }

        private void bttnOpen_MouseLeave(object sender, EventArgs e)
        {
         
            bttnOpen.BackColor = Color.FromArgb(30, 50, 90);
        }


        private void bttnExport_MouseLeave(object sender, EventArgs e)
        {
            bttnExport.BackColor = Color.FromArgb(30, 50, 90);
        }

        private void bttnStop_MouseLeave(object sender, EventArgs e)
        {
            bttnStop.BackColor = Color.FromArgb(30, 50, 90);
        }

        private void bttnClose_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void bttnClose_MouseEnter(object sender, EventArgs e)
        {         
            bttnClose.BackColor = Color.FromArgb(30, 84, 161);
        }

        private void bttnClose_MouseLeave(object sender, EventArgs e)
        {
            bttnClose.BackColor = Color.FromArgb(30, 50, 90);
        }

        [DllImport("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
        [DllImport("user32.dll")]
        public static extern bool ReleaseCapture();
        private void panel1_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ReleaseCapture();
                SendMessage(Handle, buttondown, HT_CAPTION, 0);
            }
            
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            WindowState = FormWindowState.Minimized;
        }

        private void button1_MouseEnter(object sender, EventArgs e)
        {
            button1.BackColor = Color.FromArgb(30, 84, 161);
        }

        private void button1_MouseLeave(object sender, EventArgs e)
        {
            button1.BackColor = Color.FromArgb(30, 50, 90);
        }
    }
}

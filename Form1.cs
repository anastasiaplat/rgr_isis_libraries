using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using ADODB;
using Microsoft.Office.Core;

namespace rgr1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            Excel.Application appE = new Excel.Application();
            appE.Visible = true;
            Excel.Workbook wb1 = appE.Workbooks.Open(@"D:\учеба\4 курс\исис\data.xlsx");
            Excel.Workbook wb2 = appE.Workbooks.Add();
            wb2.Worksheets[1].Cells[1, 1].Value = "Фамилии";
            wb2.Worksheets[1].Cells[1, 2].Value = "Стипендия";
            int i = 2;
            while (wb1.Worksheets[1].Cells[i, 1].Value != null)
            {
                Students student = new Students();
                student.Name = (string)wb1.Worksheets[1].Cells[i, 1].Value;
                student.Mat = (int)wb1.Worksheets[1].Cells[i, 2].Value;
                student.Fiz = (int)wb1.Worksheets[1].Cells[i, 3].Value;
                student.Him = (int)wb1.Worksheets[1].Cells[i, 4].Value;
                student.Rus = (int)wb1.Worksheets[1].Cells[i, 5].Value;
                student.Work = (string)wb1.Worksheets[1].Cells[i, 6].Value;
                wb2.Worksheets[1].Cells[i, 1].Value = student.Name;
                wb2.Worksheets[1].Cells[i, 2].Value = student.Stipa();
                i++;
            }
            wb1.Close();
            wb2.SaveAs(@"D:\учеба\4 курс\исис\excel-excel_tabl.xlsx");
            appE.Workbooks["excel-excel_tabl.xlsx"].Close();
            appE.Quit();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Excel.Application appE = new Excel.Application();
            appE.Visible = true;
            Excel.Workbook wb1 = appE.Workbooks.Open(@"D:\учеба\4 курс\исис\data.xlsx");
            Excel.Workbook wb2 = appE.Workbooks.Add();
            wb2.Worksheets[1].Cells[1, 1].Value = "Фамилии";
            wb2.Worksheets[1].Cells[1, 2].Value = "Стипендия";
            int i = 2;
            while (wb1.Worksheets[1].Cells[i, 1].Value != null)
            {
                Students student = new Students();
                student.Name = (string)wb1.Worksheets[1].Cells[i, 1].Value;
                student.Mat = (int)wb1.Worksheets[1].Cells[i, 2].Value;
                student.Fiz = (int)wb1.Worksheets[1].Cells[i, 3].Value;
                student.Him = (int)wb1.Worksheets[1].Cells[i, 4].Value;
                student.Rus = (int)wb1.Worksheets[1].Cells[i, 5].Value;
                student.Work = (string)wb1.Worksheets[1].Cells[i, 6].Value;
                wb2.Worksheets[1].Cells[i, 1].Value = student.Name;
                wb2.Worksheets[1].Cells[i, 2].Value = student.Stipa();
                i++;
            }
            Excel.Chart ch = appE.Charts.Add();
            i--;
            string s = "B2:e" + i.ToString();
            Excel.Range range = wb1.Worksheets[1].Range(s);
            ch.SetSourceData(range, Type.Missing);
            ch.ChartType = Excel.XlChartType.xl3DColumn;
            ch.HasDataTable = true;
            ch.DataTable.Font.Size = 14;
            ch.HasTitle = true;
            ch.ChartTitle.Text = "Стипендии";
            ch.ChartTitle.Font.Size = 24;
            ch.ChartTitle.Font.Color = 200;
            for (int j = 1; j < 4; j++)
                ch.Legend.LegendEntries(j).Font.Size = 12;
            Excel.Axis ox = ch.Axes(Excel.XlAxisType.xlCategory);
            ox.HasTitle = false;

            Excel.Axis oy = ch.Axes(Excel.XlAxisType.xlSeriesAxis);
            oy.HasTitle = true;
            oy.AxisTitle.Text = "Предметы";

            Excel.Axis oz = ch.Axes(Excel.XlAxisType.xlValue);
            oz.HasTitle = true;
            oz.AxisTitle.Text = "Оценки";

            ox.HasMajorGridlines = true;
            oy.HasMajorGridlines = true;
            oz.MajorUnit = 1;

            Excel.SeriesCollection ser = ch.SeriesCollection();
            ser.Item(1).Name = "Математика";
            ser.Item(2).Name = "Физика";
            ser.Item(3).Name = "Химия";
            ser.Item(4).Name = "Русский";

            string v = "={\""; for (int j = 2; j < i; j++)
                v += wb1.Worksheets[1].Cells[j, 1].Value + "\";\"";
            v += wb1.Worksheets[1].Cells[i, 1].Value + "\"}";
            ser.Item(1).XValues = v;
            wb1.Close();

            Excel.Shape sh = wb2.Worksheets[1].Shapes.
                            AddChart(Type.Missing, 10, 180, 400, 300);
            ch = sh.Chart;

            s = "b2:b" + i.ToString();
            range = wb2.Worksheets[1].Range(s);
            ch.SetSourceData(range);
            ch.HasTitle = true;
            ch.ChartTitle.Text = "Сессия";
            ch.ChartTitle.Font.Size = 24;
            ch.ChartTitle.Font.Color = 100;

            ch.Legend.LegendEntries(1).Font.Size = 12;
            ox = ch.Axes(Excel.XlAxisType.xlCategory);
            ox.HasTitle = true;
            ox.AxisTitle.Text = "Студенты";

            oz = ch.Axes(Excel.XlAxisType.xlValue);
            oz.HasTitle = true;
            oz.AxisTitle.Text = "Стипендии";

            ox.HasMajorGridlines = true;
            oz.HasMinorGridlines = false;

            ser = ch.SeriesCollection();
            ser.Item(1).Name = "Стипендия";
            ser.Item(1).XValues = v;

            wb2.SaveAs(@"D:\учеба\4 курс\исис\excel-excel_diag.xlsx");
            appE.Workbooks["excel-excel_diag.xlsx"].Close();
            appE.Quit();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Word.Application appW = new Word.Application();
            appW.Visible = true;
            Word.Document d = appW.Documents.Add(Type.Missing);

            d.Paragraphs[1].Format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            Word.Range r = d.Paragraphs[1].Range;
            r.Text = "Стипендия";
            r.Font.Size = 16;
            r.Bold = 1;
            r.Underline = Word.WdUnderline.wdUnderlineSingle;
            r.Font.Name = "Times New Roman";

            d.Paragraphs.Add();
            d.Paragraphs[2].Format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            d.Paragraphs[2].Range.Text = "";

            Excel.Application appE = new Excel.Application();
            appE.Visible = false;
            Excel.Workbook wb = appE.Workbooks.Open(@"D:\учеба\4 курс\исис\data.xlsx");

            int i = 2;
            string s;
            while (wb.Worksheets[1].Cells[i, 1].Value != null)
            {
                d.Paragraphs.Add();
                d.Paragraphs[i + 1].Format.Alignment =
                    Word.WdParagraphAlignment.wdAlignParagraphLeft;
                r = d.Paragraphs[i + 1].Range;
                r.Italic = 1;
                r.Font.SmallCaps = 0;
                r.Underline = Word.WdUnderline.wdUnderlineNone;
                r.Font.Size = 14;
                r.Bold = 0;
                r.Font.Name = "Times New Roman";
                Students student = new Students();
                student.Name = (string)wb.Worksheets[1].Cells[i, 1].Value;
                student.Mat = (int)wb.Worksheets[1].Cells[i, 2].Value;
                student.Fiz = (int)wb.Worksheets[1].Cells[i, 3].Value;
                student.Him = (int)wb.Worksheets[1].Cells[i, 4].Value;
                student.Rus = (int)wb.Worksheets[1].Cells[i, 5].Value;
                student.Work = (string)wb.Worksheets[1].Cells[i, 6].Value;
                r.Text = wb.Worksheets[1].Cells[i, 1].Value + "      " + student.Stipa().ToString();
                i++;
            }
            wb.Close();
            appE.Quit();
            d.SaveAs(@"D:\учеба\4 курс\исис\excel-word_list.docx");
            d.Close();
            appW.Quit();
        }
        private void button4_Click(object sender, EventArgs e)
        {
            Word.Application appW = new Word.Application();
            appW.Visible = true;
            Word.Document d = appW.Documents.Add(Type.Missing);
            d.Paragraphs[1].Format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            Word.Range r = d.Paragraphs[1].Range;
            r.Text = "Итоги сессии";
            r.Font.Size = 16;
            r.Bold = 1;
            r.Font.Name = "Times New Roman";
            d.Paragraphs.Add();
            r = d.Paragraphs[2].Range;
            Word.Table t = d.Tables.Add(r, 1, 2, Word.WdTableFieldSeparator.wdSeparateByCommas);

            r = t.Cell(1, 1).Range;
            r.Font.Size = 14;
            r.Text = "Фамилия";

            r = t.Cell(1, 2).Range;
            r.Font.Size = 14;
            r.Text = "Стипендия";

            Excel.Application appE = new Excel.Application();
            appE.Visible = false;
            Excel.Workbook wb = appE.Workbooks.Open(@"D:\учеба\4 курс\исис\data.xlsx");

            int i = 2;
            string s;
            while (wb.Worksheets[1].Cells[i, 1].Value != null)
            {
                t.Rows.Add();
                r = t.Cell(i, 1).Range;
                r.Bold = 0;
                r.Font.Name = "Times New Roman";
                r.Text = wb.Worksheets[1].Cells[i, 1].Value;
                r = t.Cell(i, 2).Range;
                r.Font.Name = "Times New Roman";
                Students student = new Students();
                student.Name = (string)wb.Worksheets[1].Cells[i, 1].Value;
                student.Mat = (int)wb.Worksheets[1].Cells[i, 2].Value;
                student.Fiz = (int)wb.Worksheets[1].Cells[i, 3].Value;
                student.Him = (int)wb.Worksheets[1].Cells[i, 4].Value;
                student.Rus = (int)wb.Worksheets[1].Cells[i, 5].Value;
                student.Work = (string)wb.Worksheets[1].Cells[i, 6].Value;
                t.Cell(i, 1).Range.Text = student.Name;
                t.Cell(i, 2).Range.Text = student.Stipa().ToString();
                i++;
            }
            wb.Close();
            appE.Quit();
            d.SaveAs(@"D:\учеба\4 курс\исис\excel-word_tabl.docx");
            d.Close();
            appW.Quit();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            PowerPoint.Application appP = new PowerPoint.Application();
            appP.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
            PowerPoint.Presentation p = appP.Presentations.Add();
            Excel.Application appE = new Excel.Application();
            appE.Visible = false;
            Excel.Workbook wb = appE.Workbooks.Open(@"D:\учеба\4 курс\исис\data.xlsx");
            PowerPoint.Slide sl = p.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutTitle);
            sl.Shapes[1].TextFrame.TextRange.Text = "Итоги сессии";
            sl.Shapes[2].TextFrame.TextRange.Text = "Платонова";
            PowerPoint.Slide sl1 = p.Slides.Add(2, PowerPoint.PpSlideLayout.ppLayoutObject);
            sl1.Shapes[1].TextFrame.TextRange.Text = "Фотография";
            sl1.Shapes.AddPicture(@"C:\Users\Анастасия\Desktop\0.jfif", MsoTriState.msoFalse, MsoTriState.msoTrue, 120, 150, 500, 330);
            PowerPoint.Slide sl2 = p.Slides.Add(3, PowerPoint.PpSlideLayout.ppLayoutText);
            sl2.Shapes[1].TextFrame.TextRange.Text = "Итоги сессии";
            int i = 2;
            string str = "";
            while (wb.Worksheets[1].Cells[i, 1].Value != null)
            {
                Students student = new Students();
                student.Name = (string)wb.Worksheets[1].Cells[i, 1].Value;
                student.Mat = (int)wb.Worksheets[1].Cells[i, 2].Value;
                student.Fiz = (int)wb.Worksheets[1].Cells[i, 3].Value;
                student.Him = (int)wb.Worksheets[1].Cells[i, 4].Value;
                student.Rus = (int)wb.Worksheets[1].Cells[i, 5].Value;
                student.Work = (string)wb.Worksheets[1].Cells[i, 6].Value;
                str += student.Name + " " + student.Stipa().ToString() + "\r";
                i++;
            }
            sl2.Shapes[2].TextFrame.TextRange.Text = str;
            sl2.Shapes[2].TextFrame.TextRange.Font.Size = 8;
            PowerPoint.Slide sl3 = p.Slides.Add(4, PowerPoint.PpSlideLayout.ppLayoutTable);
            sl3.Shapes[1].TextFrame.TextRange.Text = "Стипендия";
            sl3.Shapes.AddTable(1, 2);
            PowerPoint.Table t = sl3.Shapes[2].Table;
            t.Cell(1, 1).Shape.TextFrame.TextRange.Text = "Фамилия";
            t.Cell(1, 2).Shape.TextFrame.TextRange.Text = "Стипендия";
            int z = 2;
            while (wb.Worksheets[1].Cells[z, 1].Value != null)
            {
                t.Rows.Add();
                t.Cell(z, 1).Shape.TextFrame.TextRange.Text = wb.Worksheets[1].Cells[z, 1].Value;
                Students student = new Students((string)wb.Worksheets[1].Cells[z, 1].Value,
                    (int)wb.Worksheets[1].Cells[z, 2].Value,
                    (int)wb.Worksheets[1].Cells[z, 3].Value,
                    (int)wb.Worksheets[1].Cells[z, 4].Value,
                    (int)wb.Worksheets[1].Cells[z, 5].Value,
                    (string)wb.Worksheets[1].Cells[z, 6].Value);
                t.Cell(z, 2).Shape.TextFrame.TextRange.Text = student.Stipa().ToString();
                z++;
            }
            wb.Close();
            appE.Quit();
            p.SlideShowSettings.Run();
        }
        private void button6_Click(object sender, EventArgs e)
        {
            Word.Application appW = new Word.Application();
            appW.Visible = true;
            Word.Document d = appW.Documents.Open(@"D:\учеба\4 курс\исис\data.docx");
            int n = d.Paragraphs.Count;
            PowerPoint.Application appPp = new PowerPoint.Application();
            appPp.Visible = MsoTriState.msoTrue;
            PowerPoint.Presentation p = appPp.Presentations.Add();
            PowerPoint.Slide sl = p.Slides.Add(1,
            PowerPoint.PpSlideLayout.ppLayoutTitle);
            sl.Shapes[1].TextFrame.TextRange.Text = d.Paragraphs[1].Range.Text;
            sl.Shapes[2].TextFrame.TextRange.Text = d.Paragraphs[2].Range.Text;
            for (int k = 2; k < n; k++)
            {
                sl = p.Slides.Add(k, PowerPoint.PpSlideLayout.ppLayoutText);
                sl.Shapes[2].TextFrame.TextRange.Text = d.Paragraphs[k + 1].Range.Text;
            }
            d.Close();
            appW.Quit();
            int m = (n + 2) / 4;
            int[] ind1 = new int[m];
            for (int k = 1; k <= m; k++)
                ind1[k - 1] = 4 * k - 3;
            PowerPoint.SlideShowTransition sst1 = p.Slides.Range(ind1).SlideShowTransition;
            sst1.AdvanceOnTime = MsoTriState.msoTrue;
            sst1.AdvanceTime = 3;
            sst1.EntryEffect = PowerPoint.PpEntryEffect.ppEffectCoverLeftDown;
            m = (n + 1) / 4;
            int[] ind2 = new int[m];
            for (int k = 1; k <= m; k++)
                ind2[k - 1] = 4 * k - 2;
            PowerPoint.SlideShowTransition sst2 = p.Slides.Range(ind2).SlideShowTransition;
            sst2.AdvanceOnTime = MsoTriState.msoTrue;
            sst2.AdvanceTime = 3;
            sst2.EntryEffect = PowerPoint.PpEntryEffect.ppEffectCoverRightDown;
            m = n / 4;
            int[] ind3 = new int[m];
            for (int k = 1; k <= m; k++)
                ind3[k - 1] = 4 * k - 1;
            PowerPoint.SlideShowTransition sst3 = p.Slides.Range(ind3).SlideShowTransition;
            sst3.AdvanceOnTime = MsoTriState.msoTrue;
            sst3.AdvanceTime = 3;
            sst3.EntryEffect = PowerPoint.PpEntryEffect.ppEffectCoverLeftUp;
            m = (n - 1) / 4;
            int[] ind4 = new int[m];
            for (int k = 1; k <= m; k++)
                ind4[k - 1] = 4 * k;
            PowerPoint.SlideShowTransition sst4 = p.Slides.Range(ind4).SlideShowTransition;
            sst4.AdvanceOnTime = MsoTriState.msoTrue;
            sst4.AdvanceTime = 3;
            sst4.EntryEffect = PowerPoint.PpEntryEffect.ppEffectCoverRightUp;
            p.SlideShowSettings.Run();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            ADODB.Connection cnn = new Connection();
            cnn.Open(@"Provider=Microsoft.ACE.OLEDB.12.0; Data Source=D:\учеба\4 курс\исис\Database1.accdb");
            ADODB.Recordset rst = new ADODB.Recordset();
            rst.Open("Select * From Table1", cnn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic);
            rst.MoveFirst();
            do
            {
                Students student = new Students();
                student.Name = rst.Fields[1].Value;
                student.Mat = (int)rst.Fields[2].Value;
                student.Fiz = (int)rst.Fields[3].Value;
                student.Him = (int)rst.Fields[4].Value;
                student.Rus = (int)rst.Fields[5].Value;
                student.Work = rst.Fields[6].Value;
                rst.Fields["Стипендия"].Value = student.Stipa();
                rst.Move(1);
            }
            while (!rst.EOF);
            rst.Close();
            cnn.Close();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            Word.Application appWord = new Word.Application();
            Word.Document d = appWord.Documents.Add();
            appWord.Visible = true;
            ADODB.Connection cnn = new Connection();
            cnn.Open(@"Provider=Microsoft.ACE.OLEDB.12.0; Data Source=D:\учеба\4 курс\исис\Database1.accdb");
            ADODB.Recordset rst = new ADODB.Recordset();
            rst.Open("Select * From Table1", cnn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic);
            d.Paragraphs[1].Format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            Word.Range r = d.Paragraphs[1].Range;
            r.Text = "Стипендия";
            r.Font.Size = 24;
            r.Bold = 1;
            d.Paragraphs.Add();
            d.Paragraphs[2].Format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            d.Paragraphs[2].Range.Text = "";
            int i = 2;
            rst.MoveFirst();
            do
            {
                Students student = new Students();
                student.Name = rst.Fields[1].Value;
                student.Mat = (int)rst.Fields[2].Value;
                student.Fiz = (int)rst.Fields[3].Value;
                student.Him = (int)rst.Fields[4].Value;
                student.Rus = (int)rst.Fields[5].Value;
                student.Work = rst.Fields[6].Value;
                d.Paragraphs.Add();
                d.Paragraphs[i].Format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                r = d.Paragraphs[i].Range;
                r.Font.Size = 16;
                r.Bold = 0;
                d.Paragraphs[i].Range.Text = rst.Fields[1].Value + " " + student.Stipa().ToString();
                i++;
                rst.Move(1);
            }
            while (!rst.EOF);
            rst.Close();
            cnn.Close();
        }
    }
}
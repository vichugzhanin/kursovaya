# kursovaya
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word= Microsoft.Office.Interop.Word;
namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        #region Глобальные переменные
        private Word.Application wordapp; //приложение
        private Word.Documents worddocuments; //документ
        private Word.Document worddocument;
        private object ooMissing; //вспомогательная
        private object unit; //вспомогательная
        private object extend; //вспомогательная
        private int f = 0; //счётчик параграфов
        private int countInsertBreak = 0; //счётчик, по которому определяется надо ли вставлять разрыв (значения: 1 - разрыв был вставлен, новый разрыв не нужен; 0 - вставляем разрыв)
        public bool job = false; // вентиль, по которому определяется был ли создан документ
        #endregion
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            int i = Convert.ToInt32(((Button)(sender)).Tag);
            switch (i)
            {
                case 1:
                    try
                    {
                        #region Создание_документа
                        bool cont = true;
                        foreach (Control t in this.Controls)  //проверка заполненности полей
                        { if (t is TextBox)
                                if (t.Text == "") cont = false;
                        }
                        if (!cont) MessageBox.Show("Необходимо заполнить все поля");
                                else
                                {
                                    wordapp = new Word.Application();
                                    wordapp.Visible = true;
                                    Object template = Type.Missing; //переменные для создания копии Word.App
                                    Object newTemplate = false;
                                    Object documentType = Word.WdNewDocumentType.wdNewBlankDocument;
                                    Object visible = true;
                                    object oMissing = System.Reflection.Missing.Value;
                                    worddocument =
                                    wordapp.Documents.Add(
                                    ref template, ref newTemplate, ref documentType, ref visible);
                                    worddocuments = wordapp.Documents;
                                    Object name = "Документ1";
                                    worddocument = (Word.Document)worddocuments.get_Item(ref name);
                                    worddocument.Activate(); //активация приложения
                                    job = true;
                                    //Курсор ввода устанавливается в начало документа
                                    unit = Word.WdUnits.wdStory;
                                    extend = Word.WdMovementType.wdMove;
                                    wordapp.Selection.HomeKey(ref unit, ref extend);
                                }
                        #endregion
                        #region Титульник
                        string[] string1 = new string[7] 
                        { "МИНИСТЕРСТВО ОБРАЗОВАНИЯ И НАУКИ\n", "ФЕДЕРАЛЬНОЕ ГОСУДАРСТВЕННОЕ БЮДЖЕТНОЕ\n", 
                            "ОБРАЗОВАТЕЛЬНОЕ УЧРЕЖДЕНИЕ\n", "ВЫСШЕГО ПРОФЕССИОНАЛЬНОГО ОБРАЗОВАНИЯ\n", 
                            "ВЯТСКИЙ ГОСУДАРСТВЕННЫЙ УНИВЕРСИТЕТ\n", "ФАКУЛЬТЕТ ЭКОНОМИКИ И МЭНЕДЖМЕНТА\n", "КАФЕДРА " };
                        for (int ii = 0; ii < 45; ii++)
                        {
                            if (ii < 7)
                            {
                                if (ii == 6)
                                {
                                    worddocument.Paragraphs[ii + 1].Range.Text = string1[ii] + textBox8.Text.ToUpper(); //+название кафедры
                                    worddocument.Paragraphs[ii + 1].Range.Font.Size = 14; //размер шрифта
                                    worddocument.Paragraphs[ii + 1].Range.Font.Name = "Times New Roman"; //шрифт
                                    worddocument.Paragraphs[ii + 1].Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                                    worddocument.Paragraphs[ii + 1].Range.ParagraphFormat.LeftIndent =
                                    worddocument.Paragraphs[ii + 1].Range.Application.CentimetersToPoints(0); //отступ слева
                                    worddocument.Paragraphs[ii + 1].Range.ParagraphFormat.RightIndent =
                                     worddocument.Paragraphs[ii + 1].Range.Application.CentimetersToPoints(0); //отступ справа
                                    worddocument.Paragraphs[ii + 1].Range.ParagraphFormat.LineSpacing = (float)12; //межстрочный интервал
                                    worddocument.Paragraphs[ii + 1].Range.ParagraphFormat.LineUnitAfter
                                        = worddocument.Paragraphs[ii + 1].Range.Application.PointsToInches((float)1); //интервал ПОСЛЕ
                                }
                                else
                                {
                                    worddocument.Paragraphs[ii + 1].Range.Text = string1[ii];
                                    worddocument.Paragraphs[ii + 1].Range.Font.Size = 14; //размер шрифта
                                    worddocument.Paragraphs[ii + 1].Range.Font.Name = "Times New Roman"; //шрифт
                                    worddocument.Paragraphs[ii + 1].Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                                    worddocument.Paragraphs[ii + 1].Range.ParagraphFormat.LeftIndent =
                                    worddocument.Paragraphs[ii + 1].Range.Application.CentimetersToPoints(0); //отступ слева
                                    worddocument.Paragraphs[ii + 1].Range.ParagraphFormat.RightIndent =
                                     worddocument.Paragraphs[ii + 1].Range.Application.CentimetersToPoints(0); //отступ справа
                                    worddocument.Paragraphs[ii + 1].Range.ParagraphFormat.LineSpacing = (float)12; //межстрочный интервал
                                    worddocument.Paragraphs[ii + 1].Range.ParagraphFormat.LineUnitAfter
                                        = worddocument.Paragraphs[ii + 1].Range.Application.PointsToInches((float)1); //интервал ПОСЛЕ
                                }

                                
                            }
                            else if (ii>=7 && ii<=12) 
                            {
                                ooMissing = System.Reflection.Missing.Value; //текущий номер параграфа
                                worddocument.Paragraphs.Add(ref ooMissing); //добавляем параграф
                                worddocument.Paragraphs[ii + 1].Range.Text = "\n"; //вписываем \n

                            }
                            else if (ii==13)
                            {
                                if (textBox1.Lines.Length > 1)
                                {
                                    worddocument.Paragraphs[ii + 1].Range.Text = textBox1.Text;
                                    for (int iii=0; iii<textBox1.Lines.Length; iii++)
                                    {
                                    worddocument.Paragraphs[iii + ii + 1].Range.Font.Size = 14; //размер шрифта
                                    worddocument.Paragraphs[iii + ii + 1].Range.Font.Bold = 1;
                                    worddocument.Paragraphs[iii + ii + 1].Range.Font.Name = "Times New Roman"; //шрифт
                                    worddocument.Paragraphs[iii + ii + 1].Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                                    worddocument.Paragraphs[iii + ii + 1].Range.ParagraphFormat.LeftIndent =
                                    worddocument.Paragraphs[iii + ii + 1].Range.Application.CentimetersToPoints(0); //отступ слева
                                    worddocument.Paragraphs[iii + ii + 1].Range.ParagraphFormat.RightIndent =
                                     worddocument.Paragraphs[iii + ii + 1].Range.Application.CentimetersToPoints(0); //отступ справа
                                    worddocument.Paragraphs[iii + ii + 1].Range.ParagraphFormat.LineSpacing = (float)12; //межстрочный интервал
                                    worddocument.Paragraphs[iii + ii + 1].Range.ParagraphFormat.LineUnitAfter
                                        = worddocument.Paragraphs[iii + ii + 1].Range.Application.PointsToInches((float)1); //интервал ПОСЛЕ
                                    }
                                    f += textBox1.Lines.Length; ii += textBox1.Lines.Length;
                                }
                                else
                                {
                                    worddocument.Paragraphs[ii + 1].Range.Text = textBox1.Text;
                                    worddocument.Paragraphs[ii + 1].Range.Font.Size = 14; //размер шрифта
                                    worddocument.Paragraphs[ii + 1].Range.Font.Bold = 1;
                                    worddocument.Paragraphs[ii + 1].Range.Font.Name = "Times New Roman"; //шрифт
                                    worddocument.Paragraphs[ii + 1].Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                                    worddocument.Paragraphs[ii + 1].Range.ParagraphFormat.LeftIndent =
                                    worddocument.Paragraphs[ii + 1].Range.Application.CentimetersToPoints(0); //отступ слева
                                    worddocument.Paragraphs[ii + 1].Range.ParagraphFormat.RightIndent =
                                     worddocument.Paragraphs[ii + 1].Range.Application.CentimetersToPoints(0); //отступ справа
                                    worddocument.Paragraphs[ii + 1].Range.ParagraphFormat.LineSpacing = (float)12; //межстрочный интервал
                                    worddocument.Paragraphs[ii + 1].Range.ParagraphFormat.LineUnitAfter
                                        = worddocument.Paragraphs[ii + 1].Range.Application.PointsToInches((float)1); //интервал ПОСЛЕ
                                }

                                
                            }
                            else if (ii == 15 + textBox1.Lines.Length)
                            {
                                ooMissing = System.Reflection.Missing.Value;
                                worddocument.Paragraphs.Add(ref ooMissing);
                                worddocument.Paragraphs[ii - 1].Range.Text = "\n";
                                worddocument.Paragraphs[ii].Range.Text = "Курсовая работа";
                                worddocument.Paragraphs[ii].Range.Font.Bold = 0;
                                for (int j=0; j<7; j++) //перевод каретки х7
                                {
                                    ooMissing = System.Reflection.Missing.Value;
                                    worddocument.Paragraphs.Add(ref ooMissing);
                                    worddocument.Paragraphs[ii + j + 1].Range.Text = "\n";
                                    ii++;
                                }
                                worddocument.Paragraphs[ii + 1].Range.Text = "студента " + textBox3.Text + " курса " + textBox4.Text+"\n";
                                worddocument.Paragraphs[ii + 1].Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight; ii++;
                                worddocument.Paragraphs[ii + 1].Range.Text = "группа " + textBox5.Text + "\n";
                                worddocument.Paragraphs[ii + 1].Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight; ii++;
                                worddocument.Paragraphs[ii + 1].Range.Text = textBox6.Text + "\n";
                                worddocument.Paragraphs[ii + 1].Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight; ii++;
                                worddocument.Paragraphs[ii + 1].Range.Text = "\n\n"; ii += 2;
                                worddocument.Paragraphs[ii + 1].Range.Text = "Руководитель\n";
                                worddocument.Paragraphs[ii + 1].Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight; ii++;
                                worddocument.Paragraphs[ii + 1].Range.Text = textBox7.Text + "\n\n\n";
                                worddocument.Paragraphs[ii + 1].Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight; ii+=3;
                                worddocument.Paragraphs[ii + 1].Range.Text = "Дата защиты: «__» _______________ 20__ г.\n\n\n";
                                worddocument.Paragraphs[ii + 1].Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight; ii += 3;
                                worddocument.Paragraphs[ii + 1].Range.Text = "Оценка: _____________________\n\n\n\n\n\n";
                                worddocument.Paragraphs[ii + 1].Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight; ii += 6;
                                worddocument.Paragraphs[ii + 1].Range.Text = "Киров\n";
                                worddocument.Paragraphs[ii + 1].Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter; ii++;
                                worddocument.Paragraphs[ii + 1].Range.Text = textBox9.Text;
                                worddocument.Paragraphs[ii + 1].Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                                worddocument.Paragraphs[ii + 1].Range.ParagraphFormat.LineUnitAfter = 0; //интервал после 
                                ii++;
                            }
                            f = ii;
                        }
                        #endregion
                        //вставка разрыва на след. страницу
                        #region Разрыв
                        unit = Word.WdUnits.wdStory;
                        extend = Word.WdMovementType.wdMove;
                        wordapp.Selection.EndKey(ref unit, ref extend);
                        object oType;
                        oType = Word.WdBreakType.wdSectionBreakNextPage;
                        //И на новый лист
                        wordapp.Selection.InsertBreak(ref oType);
                        countInsertBreak++;
                        #endregion
                        #region Анализ текста для содержания
                        ooMissing = System.Reflection.Missing.Value;
                        f = worddocument.Paragraphs.Count;
                        worddocument.Paragraphs.Add(ref ooMissing);
                        worddocument.Paragraphs[f].Range.Text = "Содержание\n\n";
                        countInsertBreak = 0;
                        worddocument.Paragraphs[f].Range.ParagraphFormat.LineSpacing = (float)13.8; //межстрочный интервал
                        worddocument.Paragraphs[f].Range.ParagraphFormat.LineUnitAfter
                            = worddocument.Paragraphs[f].Range.Application.PointsToInches((float)8);
                        f ++;
                        worddocument.Paragraphs[f + 1].Range.Text = "Введение..............................\n";
                        worddocument.Paragraphs[f+1].Range.ParagraphFormat.LineUnitAfter
                            = worddocument.Paragraphs[f+1].Range.Application.PointsToInches((float)8);
                        worddocument.Paragraphs[f + 1].Range.ParagraphFormat.LineSpacing = (float)13.8;
                        worddocument.Paragraphs[f + 1].Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft; f++;
                        string[] glava =  new string[]{"1.","1.1","1.2","1.3","1.4","1.5","1.6","1.7","1.8","1.9",
                            "2.","2.1","2.2","2.3","2.4","2.5","2.6","2.7","2.8","2.9"};
                        for (int j=0;j<textBox2.Lines.Length;j++)
                        {
                            for(int ch=0;ch<glava.Length;ch++)
                            {
                                if (textBox2.Lines[j] == glava[ch])  //проверка совпадений
                                {
                                worddocument.Paragraphs[f + 1].Range.Text
                                    = textBox2.Lines[j] + "....................................\n";
                                worddocument.Paragraphs[f + 1].Range.ParagraphFormat.LineUnitAfter
                             = worddocument.Paragraphs[f + 1].Range.Application.PointsToInches((float)8);
                                worddocument.Paragraphs[f + 1].Range.ParagraphFormat.LineSpacing = (float)13.8;
                                worddocument.Paragraphs[f + 1].Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                                f++;
                                }
                            }
                        }
                        worddocument.Paragraphs[f + 1].Range.Text = "Заключение............................\n";
                        worddocument.Paragraphs[f + 1].Range.ParagraphFormat.LineUnitAfter
                            = worddocument.Paragraphs[f + 1].Range.Application.PointsToInches((float)8);
                        worddocument.Paragraphs[f + 1].Range.ParagraphFormat.LineSpacing = (float)13.8;
                        worddocument.Paragraphs[f + 1].Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                        f++;
                        worddocument.Paragraphs[f + 1].Range.Text = "Список литературы.....................\n";
                        worddocument.Paragraphs[f + 1].Range.ParagraphFormat.LineUnitAfter
                            = worddocument.Paragraphs[f + 1].Range.Application.PointsToInches((float)8);
                        worddocument.Paragraphs[f + 1].Range.ParagraphFormat.LineSpacing = (float)13.8;
                        worddocument.Paragraphs[f + 1].Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                        f++;
                        worddocument.Paragraphs[f + 1].Range.Text = "Рецензия............................\n";
                        worddocument.Paragraphs[f + 1].Range.ParagraphFormat.LineUnitAfter
                            = worddocument.Paragraphs[f + 1].Range.Application.PointsToInches((float)8);
                        worddocument.Paragraphs[f + 1].Range.ParagraphFormat.LineSpacing = (float)13.8;
                        worddocument.Paragraphs[f + 1].Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                        f++;
                        #endregion
                        #region Разрыв
                        unit = Word.WdUnits.wdStory;
                        extend = Word.WdMovementType.wdMove;
                        wordapp.Selection.EndKey(ref unit, ref extend);
                        oType = Word.WdBreakType.wdSectionBreakNextPage;
                        //И на новый лист
                        wordapp.Selection.InsertBreak(ref oType);
                        countInsertBreak++;
                        #endregion
                        #region Анализ текста для основного текста
                        string[] glava1 = new string[] { "Введение","Заключение","Список литературы", "Рецензия",
                            "1.", "1.1", "1.2", "1.3", "1.4", "1.5", "1.6", "1.7", "1.8", "1.9", 
                            "2.", "2.1", "2.2", "2.3", "2.4", "2.5", "2.6", "2.7", "2.8", "2.9" };
                        for (int j = 0; j < textBox2.Lines.Length; j++)
                        {
                                if (glava1.Contains(textBox2.Lines[j]))
                                {
                                    if (countInsertBreak==0) //проверка на наличие разрыва ДО
                                    {
                                        #region Разрыв
                                        unit = Word.WdUnits.wdStory;
                                        extend = Word.WdMovementType.wdMove;
                                        wordapp.Selection.EndKey(ref unit, ref extend);
                                        oType = Word.WdBreakType.wdSectionBreakNextPage;
                                        //И на новый лист
                                        wordapp.Selection.InsertBreak(ref oType);
                                        countInsertBreak++;
                                        #endregion
                                    }
                                    
                                    f = worddocument.Paragraphs.Count;
                                    worddocument.Paragraphs[f].Range.Text
                                        = textBox2.Lines[j]+"\n\n";
                                    countInsertBreak = 0;
                                    worddocument.Paragraphs[f].Range.ParagraphFormat.LineUnitAfter
                                 = worddocument.Paragraphs[f].Range.Application.PointsToInches((float)8);
                                    worddocument.Paragraphs[f].Range.ParagraphFormat.LineSpacing = (float)13.8;
                                    worddocument.Paragraphs[f].Range.Bold = 1;
                                    worddocument.Paragraphs[f].Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                                    f+=2;
                                }
                                else
                                {
                                    worddocument.Paragraphs[f].Range.Text
                                        = textBox2.Lines[j] + "\n";
                                    countInsertBreak = 0;
                                    worddocument.Paragraphs[f].Range.ParagraphFormat.LineUnitAfter
                                 = worddocument.Paragraphs[f].Range.Application.PointsToInches((float)8);
                                    worddocument.Paragraphs[f].Range.ParagraphFormat.LineSpacing = (float)13.8;
                                    worddocument.Paragraphs[f].Range.Bold = 0;
                                    worddocument.Paragraphs[f].Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify;
                                    f++;
                                }
                            }
                        #endregion
                        InsertPageNumbers(worddocument, Word.WdPageNumberAlignment.wdAlignPageNumberCenter);

                    }
                    catch (Exception ex)
                    {
                        Text = ex.Message;
                    }
                    break;
                case 2:
                    
                    break;
                case 3:
                    Object saveChanges = Word.WdSaveOptions.wdDoNotSaveChanges;
                    Object originalFormat = Word.WdOriginalFormat.wdWordDocument;
                    Object routeDocument = Type.Missing;
                    wordapp.Quit(ref saveChanges, ref originalFormat, ref routeDocument);
                    wordapp = null;
                    break;
                default:
                    Application.Exit();
                    break;
            }
        }

        public void InsertPageNumbers(Word.Document doc, Word.WdPageNumberAlignment alignment) //для нумерации с 3го листа
        {

                //Переход на вторую страницу (вернее, в начало третьей)
                Word.Range range = doc.Range().GoTo(Word.WdGoToItem.wdGoToPage, Word.WdGoToDirection.wdGoToAbsolute, 3);
                //Колонтитул второго раздела
                Word.HeaderFooter hf = doc.Sections[3].Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];
                //Открепление нумерации от колонтитула предыдущего раздела
                hf.LinkToPrevious = false;
                //Не начинать нумерацию с 1
                hf.PageNumbers.RestartNumberingAtSection = false;
                //Добавление нумерации по заданному выравниванию
                hf.PageNumbers.Add(alignment, true);
                hf.Range.Font.Name = "Times New Roman";
                hf.Range.Font.Size = 14;
            
        }

        public void vInsertNumberPages(int viWhere, bool bPageFirst)  //для нумерации с первого листа
        {
            object alignment = Word.WdPageNumberAlignment.wdAlignPageNumberCenter;
            object bFirstPage = bPageFirst;
            object bF = true;
            // создаём коллонтитулы 
            worddocument.ActiveWindow.ActivePane.View.SeekView = Word.WdSeekView.wdSeekCurrentPageFooter;

            switch (viWhere)
            {
                case 1:
                    alignment = Word.WdPageNumberAlignment.wdAlignPageNumberRight;
                    break;
                case 2:
                    alignment = Word.WdPageNumberAlignment.wdAlignPageNumberLeft;
                    break;
            }
            object bp = 3;
            //wordapp.Selection.HeaderFooter.Range.MoveStart(3);
            wordapp.Selection.HeaderFooter.PageNumbers.Add(ref   alignment, ref  bFirstPage);
            //worddocument.Sections[1].Footers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.Text = "text";
            // wordapp.Selection.Sections[2].Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text = "";
        }

        private void button4_Click(object sender, EventArgs e) //кнопка сохранения и выхода
        {
            //Подготавливаем параметры для сохранения документа
            string saveNow = DateTime.Now.ToString().Replace('.', ' ').Replace(':', ' ');  //формирование пригодной для названия документа строки из текущей даты
            Object fileName = Application.StartupPath+"/kurs"+saveNow+".docx";
            Object fileFormat = Word.WdSaveFormat.wdFormatDocumentDefault;
            Object lockComments = false;
            Object password = "";
            Object addToRecentFiles = false;
            Object writePassword = "";
            Object readOnlyRecommended = false;
            Object embedTrueTypeFonts = false;
            Object saveNativePictureFormat = false;
            Object saveFormsData = false;
            Object saveAsAOCELetter = Type.Missing;
            Object encoding = Type.Missing;
            Object insertLineBreaks = Type.Missing;
            Object allowSubstitutions = Type.Missing;
            Object lineEnding = Type.Missing;
            Object addBiDiMarks = Type.Missing;
            worddocument.SaveAs(ref fileName,
                ref fileFormat, ref lockComments,
                ref password, ref addToRecentFiles, ref writePassword,
                ref readOnlyRecommended, ref embedTrueTypeFonts,
                ref saveNativePictureFormat, ref saveFormsData,
                ref saveAsAOCELetter, ref encoding, ref insertLineBreaks,
                ref allowSubstitutions, ref lineEnding, ref addBiDiMarks);
            Object saveChanges = Word.WdSaveOptions.wdDoNotSaveChanges;
            Object originalFormat = Word.WdOriginalFormat.wdWordDocument;
            Object routeDocument = Type.Missing;
            wordapp.Quit(ref saveChanges, ref originalFormat, ref routeDocument);
            wordapp = null;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox9_TextChanged(object sender, EventArgs e)
        {

        }
    }
}

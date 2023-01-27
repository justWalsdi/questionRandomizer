using System;
using System.Collections.Generic;
using System.Windows;
using ClosedXML.Excel;
using Microsoft.Win32;
using Xceed.Words.NET;
using Xceed.Document.NET;

namespace QuestionRandomizer
{
    class FileTools
    {
        public static bool CreateExcelFile(List<StudentDataClass> students)
        {
            List<StudentDataClass> studentsTwo = new List<StudentDataClass>();
            foreach (var student in students)
            {
                if (student.isMarkSet)
                {
                    studentsTwo.Add(student);
                }
            }
            if (studentsTwo.Count <= 0)
            {
                MessageBox.Show("Требуется оценить хотя бы одного студента!", "Ошибка");
                return false;
            }

            string FileName = "";
            try
            {
                var saveFile = new SaveFileDialog();
                saveFile.DefaultExt = ".xlsx";
                saveFile.Filter = "Excel documents (.xlsx)|*.xlsx";

                bool? result = saveFile.ShowDialog();
                if (result == true)
                {
                    FileName = saveFile.FileName;
                }
                else
                {
                    MessageBox.Show("Место для сохранения не доступно.", "Ошибка");
                    return false;
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Нет доступа к файлу. Возможно он уже открыт?", "Ошибка");
                return false;
            }

            using var wbook = new XLWorkbook();

            var ws = wbook.Worksheets.Add("Оценки");
            try
            {
                ws.Cells($"A1:F{studentsTwo.Count+1}").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                ws.Cells($"A1:F{studentsTwo.Count+1}").Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
            } catch(Exception e)
            {
                MessageBox.Show($"{e}", "Ошибка");
            }
            ws.Cell("A1").Value = "ФИО";
            ws.Cell("B1").Value = "Лекции";
            ws.Cell("C1").Value = "Практика";
            ws.Cell("D1").Value = "Сдано практик";
            ws.Cell("E1").Value = "Экзамен";
            ws.Cell("F1").Value = "Оценка";

            
            for (int i = 0; i < studentsTwo.Count; i++)
            {
                ws.Cell($"A{i+2}").Value = studentsTwo[i].FullName;
                ws.Cell($"B{i+2}").Value = studentsTwo[i].attendAtLectures;
                ws.Cell($"C{i+2}").Value = studentsTwo[i].attendAtPractice;
                ws.Cell($"D{i+2}").Value = studentsTwo[i].tasksCompleted;
                ws.Cell($"E{i+2}").Value = studentsTwo[i].questionsAnswered;
                ws.Cell($"F{i+2}").Value = studentsTwo[i].mark;
            }
            ws.Columns("A:F").AdjustToContents();
            wbook.SaveAs(FileName);

            return true;
        }
        public static string? returnFileName()
        {
            var openFile = new OpenFileDialog();
            openFile.DefaultExt = ".txt";
            openFile.Filter = "Text documents (.txt)|*.txt";

            bool? result = openFile.ShowDialog();
            if (result == true)
            {
                return openFile.FileName;
            }
            else
            {
                return null;
            }
        }
        public static void CreateDocx(List<Questions> questions, List<StudentDataClass> students)
        {
            // Random questions for everyone
            //questions.Shuffle();
            // There will be data prepared
            string FileName;
            try
            {
                var saveFile = new SaveFileDialog();
                saveFile.DefaultExt = ".docx";
                saveFile.Filter = "Word documents (.docx)|*.docx";
                bool? result = saveFile.ShowDialog();
                if (result == true)
                {
                    FileName = saveFile.FileName;
                }
                else
                {
                    MessageBox.Show("Место для сохранения не доступно.", "Ошибка");
                    return;
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Нет доступа к файлу. Возможно он уже открыт?", "Ошибка");
                return;
            }
            DocX.Create(FileName);
            using (var doc = DocX.Create(FileName))
            {
                Dictionary<StudentDataClass, List<Questions>> studentsDict = prepareDict(students, questions);
                /* Default Values For Text Insertions */
                const string mainTitle = "Билеты по дисциплине \"Стандартизация, сертификация и управление качеством программного обеспечения\".";
                Font defaultFont = new Font("Times New Roman");
                /* Default Values For Text Insertions */

                doc.InsertParagraph(mainTitle).Font(defaultFont).FontSize(16).SpacingAfter(18).Alignment = Alignment.center;
                int twoTicketsPerSheet = 0;
                foreach (StudentDataClass student in students)
                {
                    doc.InsertParagraph($"Билет №{student.ID} {student.FullName}").Font(defaultFont).FontSize(18).SpacingAfter(6);
                    for (int i = 0; i < 3; i++)
                    {
                        doc.InsertParagraph($"Вопрос №{i + 1}\n").Bold(true).Font(defaultFont).FontSize(16).SpacingAfter(6).Alignment = Alignment.center;
                        doc.InsertParagraph($"{studentsDict[student][i].QuestionText}\n").Font(defaultFont).FontSize(15).SpacingAfter(6).Alignment = Alignment.left;
                    }
                    doc.InsertParagraph("").Font(defaultFont).FontSize(16);
                    // every two questions there is a section page break inserted
                    twoTicketsPerSheet++;
                    if (twoTicketsPerSheet > 1)
                    {
                        twoTicketsPerSheet = 0;
                        doc.InsertSectionPageBreak();
                    }
                }
                try
                {
                    doc.Save();
                    MessageBox.Show($"Файл сохранён по пути {FileName}!");
                }
                catch (Exception)
                {
                    MessageBox.Show("Нет доступа к файлу. Возможно он уже открыт?", "Ошибка");
                    return;
                }

                // while it is a good idea, can't be used because of time it takes to save a file.
                // Process.Start("explorer.exe", "/select, \"" + tempFileName + "\"");
            }
        }
        private static Dictionary<StudentDataClass, List<Questions>> prepareDict(List<StudentDataClass> students, List<Questions> questions)
        {
            Dictionary <StudentDataClass, List <Questions>> studentsDict = new Dictionary<StudentDataClass, List<Questions>>();
            int i = 0;
            if (questions[0].ID == 1)
                questions.Reverse();

            foreach (StudentDataClass student in students)
            {
                List<Questions> threeQuestions = new List<Questions>();
                for (int j = 0; j < 3; j++)
                {
                    threeQuestions.Add(questions[i]);
                    i++;
                    // if we reached the end of the questions' list
                    if (i == questions.Count)
                        i = 0;

                }
                threeQuestions.Shuffle();
                studentsDict.Add(student, threeQuestions);
            }
            return studentsDict;
        }
    }
}

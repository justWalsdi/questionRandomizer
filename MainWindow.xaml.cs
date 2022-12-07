using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using Microsoft.Win32;
using Xceed.Words.NET;
using Xceed.Document.NET;

namespace QuestionRandomizer
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        List<StudentDataClass> students = new List<StudentDataClass>();
        List<Questions> questions = new List<Questions>();
        Dictionary<StudentDataClass, List<Questions>> studentsDict = new Dictionary<StudentDataClass, List<Questions>>();
        bool isStudentListLoaded = false;
        bool isQuestionListLoaded = false;
        public MainWindow() => InitializeComponent();
        private void Window_Initialized(object sender, EventArgs e)
        {

        }
        private void StudentLoad(object sender, RoutedEventArgs e)
        {
            string? fileName = returnFileName();
            if (fileName == null)
            {
                MessageBox.Show("Файл не найден.");
                return;
            }
            string[] lines = File.ReadAllLines(fileName);
            int sID = 0;
            students.Clear();
            foreach (string line in lines)
            {
                if (line != "")
                {
                    sID++;
                    students.Add(new StudentDataClass() { ID = sID, FullName = line });
                }
            }
            StudentGrid.ItemsSource = students;
            StudentGrid.Items.Refresh();
            isStudentListLoaded = true;
            checkIfEverythingLoaded();
        }
        private void CalculateMarks(object sender, RoutedEventArgs e)
        {

        }
        private void QuestionLoad(object sender, RoutedEventArgs e)
        {
            string? fileName = returnFileName();
            if (fileName == null)
            {
                MessageBox.Show("Файл вопросов не найден.");
                return;
            }
            questions.Clear();
            string[] lines = File.ReadAllLines(fileName);
            int sID = 0;
            foreach(string line in lines)
            {
                sID++;
                if (line != "")
                    questions.Add(new Questions() { ID = sID, QuestionText = line});
            }
            QuestionGrid.ItemsSource = questions;
            QuestionGrid.Items.Refresh();
            isQuestionListLoaded = true;
            checkIfEverythingLoaded();
        }
        private string? returnFileName()
        {
            var openFile = new OpenFileDialog();
            openFile.DefaultExt = ".txt";
            openFile.Filter = "Text documents (.txt)|*.txt";

            bool? result = openFile.ShowDialog();
            if (result == true)
            {
                return openFile.FileName;
            } else
            {
                return null;
            }
        }
        private void checkIfEverythingLoaded()
        {
            if (students.Count * 3 > questions.Count && isQuestionListLoaded)
            {
                RandomCreator.Visibility = Visibility.Hidden;
                AnswerCalcTab.Visibility = Visibility.Hidden;
                studentsDict.Clear();
                MessageBox.Show("Добавьте вопросы или загрузите другой список.", "Вопросов меньше чем студентов!");
            }
            else
            {
                if (isQuestionListLoaded && isStudentListLoaded)
                {
                    RandomCreator.Visibility = Visibility.Visible;
                    AnswerCalcTab.Visibility = Visibility.Visible;
                    prepareDict();
                }
            }
        }   
        private void CreateDoc(object sender, RoutedEventArgs e)
        {
            // Random questions for everyone
            // questions.Shuffle();
            // There will be data prepared
            string fileName = Environment.ExpandEnvironmentVariables("%TEMP%\\questions.docx");
            DocX.Create(fileName);
            using (var doc = DocX.Create(fileName))
            {
                /* Default Values For Text Insertions */
                const string mainTitle = "Билеты по дисциплине \"Стандартизация, сертификация и управление качеством программного обеспечения\".";
                Font defaultFont = new Font("Times New Roman");
                /* Default Values For Text Insertions */

                doc.InsertParagraph(mainTitle).Font(defaultFont).FontSize(16).SpacingAfter(18).Alignment = Alignment.center;
                int twoTicketsPerSheet = 0;
                foreach (StudentDataClass student in students)
                {
                    doc.InsertParagraph($"Билет №{student.ID} ФИО {student.FullName}").Font(defaultFont).FontSize(14).SpacingAfter(6);
                    for (int i = 0; i < 3; i++)
                    {
                        doc.InsertParagraph($"Вопрос №{i + 1}\n").Font(defaultFont).FontSize(14).SpacingAfter(6).Alignment = Alignment.center;
                        doc.InsertParagraph($"{studentsDict[student][i].QuestionText}\n").Font(defaultFont).FontSize(12).SpacingAfter(6).Alignment = Alignment.center;
                    }
                    // every two questions there is a section page break inserted
                    twoTicketsPerSheet++;
                    if (twoTicketsPerSheet > 1)
                    {
                        twoTicketsPerSheet = 0;
                        doc.InsertSectionPageBreak();
                    }
                }
                string tempFileName;
                try
                {
                    var saveFile = new SaveFileDialog();
                    saveFile.DefaultExt = ".docx";
                    saveFile.Filter = "Word documents (.docx)|*.docx";
                    bool? result = saveFile.ShowDialog();
                    if (result == true)
                    {
                        tempFileName = saveFile.FileName;
                        doc.SaveAs(tempFileName);
                    } else
                    {
                        MessageBox.Show("Место для сохранения не доступно.","Ошибка");
                        return;
                    }
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
        private void prepareDict()
        {
            studentsDict.Clear();
            int i = 0;
            foreach (StudentDataClass student in students)
            {
                List<Questions> threeQuestions = new List<Questions>();
                for (int j = 0; j < 3; j++)
                {
                    threeQuestions.Add(questions[i]);
                    i++;
                }
                studentsDict.Add(student, threeQuestions);
            }
        }
    }
}

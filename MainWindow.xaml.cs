using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using Microsoft.Win32;
using System.Linq;
using System.Diagnostics;
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
                sID++;
                students.Add(new StudentDataClass() { ID = sID, FullName = line });
            }
            StudentGrid.ItemsSource = students;
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
            string[] lines = File.ReadAllLines(fileName);
            int sID = 0;
            questions.Clear();
            foreach(string line in lines)
            {
                sID++;
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
                MessageBox.Show("Добавьте вопросы или загрузите другой список.", "Вопросов меньше чем студентов!");
            }
            else
            {
                if (isQuestionListLoaded && isStudentListLoaded)
                {
                    RandomCreator.Visibility = Visibility.Visible;
                    AnswerCalcTab.Visibility = Visibility.Visible;
                }
            }
        }   
        private void CreateDoc(object sender, RoutedEventArgs e)
        {
            // Random questions for everyone
            questions.Shuffle();

            string fileName = Environment.ExpandEnvironmentVariables("%USERPROFILE%\\Documents\\questions.docx");

            using (var doc = DocX.Create(fileName))
            {
                /* Default Values For Text Insertions */
                const string mainTitle = "Билеты по дисциплине \"Стандартизация, сертификация и управление качеством программного обеспечения\".";
                Font defaultFont = new Font("Times New Roman");
                /* Default Values For Text Insertions */

                doc.InsertParagraph(mainTitle).Font(defaultFont).FontSize(16).SpacingAfter(18).Alignment = Alignment.center; 
                int twoTicketsPerSheet = 0;
                int questionPosition = 0;
                foreach (StudentDataClass student in students)
                {
                    doc.InsertParagraph($"Билет №{student.ID} ФИО {student.FullName}").Font(defaultFont).FontSize(14).SpacingAfter(6);
                    for (int i = 0; i < 3; i++)
                    {
                        doc.InsertParagraph($"Вопрос №{i + 1}\n").Font(defaultFont).FontSize(14).SpacingAfter(6).Alignment = Alignment.center;
                        doc.InsertParagraph($"{questions[questionPosition].QuestionText}\n").Font(defaultFont).FontSize(12).SpacingAfter(6).Alignment = Alignment.center;
                        questionPosition++;
                    }
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
                } catch(Exception error)
                {
                    MessageBox.Show("Нет доступа к файлу, возможно он уже открыт?", "Ошибка");
                    return;
                }
                MessageBox.Show($"Документ был создан здесь: {fileName}");
                Process.Start("explorer.exe", "/select, \"" + fileName + "\"");
            }
        }
    }
    static class ExtensionsClass
    {
        private static Random rng = new Random();

        public static void Shuffle<T>(this IList<T> list)
        {
            int n = list.Count;
            while (n > 1)
            {
                n--;
                int k = rng.Next(n + 1);
                T value = list[k];
                list[k] = list[n];
                list[n] = value;
            }
        }
    }
}

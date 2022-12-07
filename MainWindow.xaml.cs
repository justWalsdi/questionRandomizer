using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using Microsoft.Win32;
using Xceed.Words.NET;
using Xceed.Document.NET;
using System.Data;


namespace QuestionRandomizer
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        // Create ticket data
        List<StudentDataClass> students = new List<StudentDataClass>();
        List<Questions> questions = new List<Questions>();
        Dictionary<StudentDataClass, List<Questions>> studentsDict = new Dictionary<StudentDataClass, List<Questions>>();
        bool isStudentListLoaded = false;
        bool isQuestionListLoaded = false;
        // Create ticket data

        // Calculate marks data
        int maxAttendanceLecture = 0;
        int maxAttendancePractice = 0;
        bool isMaxAttendanceLectureSet = false;
        bool isMaxAttendancePracticeSet = false;
        private void isAnswersTabOpened()
        {
            if (isMaxAttendanceLectureSet && isMaxAttendancePracticeSet)
            {
                Answers.IsEnabled = true;
                MarksGrid.ItemsSource = students;
                MarksGrid.Items.Refresh();
                return;
            }
            Answers.IsEnabled = false;   
        }
        // Calculate marks data
        private void maxAttendanceLectureSetter(object sender, RoutedEventArgs e)
        {
            isMaxAttendanceLectureSet = false;
            isAnswersTabOpened();
            if (MaxLectureAttendTextbox.Text == null)
                return;
            if (Int32.TryParse(MaxLectureAttendTextbox.Text, out int numValue))
            {
                if (numValue < 0)
                {
                    MaxLectureAttendTextbox.Text = "0";
                    MessageBox.Show("Число не может быть меньше нуля!", "Ошибка!");
                    return;
                }
                if (numValue > 100)
                {
                    MessageBox.Show("Занятий не может быть больше, чем рабочих дней!");
                    return;
                }
                maxAttendanceLecture = numValue;
                isMaxAttendanceLectureSet = true;
                isAnswersTabOpened();
                MessageBox.Show("Количество лекционных занятий установлено!");
            } else
            {
                MaxLectureAttendTextbox.Text = "0";
                MessageBox.Show("Попробуйте ввести число", "Ошибка!");
            }
        }
        private void maxAttendancePracticeSetter(object sender, RoutedEventArgs e)
        {
            isMaxAttendancePracticeSet = false;
            isAnswersTabOpened();
            if (MaxPracticeAttendTextbox.Text == null)
                return;
            if (Int32.TryParse(MaxPracticeAttendTextbox.Text, out int numValue))
            {
                if (numValue < 0)
                {
                    MaxPracticeAttendTextbox.Text = "0";
                    MessageBox.Show("Число не может быть меньше нуля!", "Ошибка!");
                    return;
                }
                if (numValue > 100)
                {
                    MaxPracticeAttendTextbox.Text = "0";
                    MessageBox.Show("Занятий не может быть больше, чем рабочих дней!");
                    return;
                }
                maxAttendancePractice = numValue;
                isMaxAttendancePracticeSet = true;
                isAnswersTabOpened();
                MessageBox.Show("Количество практических занятий установлено!");
            }
            else
            {
                MaxPracticeAttendTextbox.Text = "0";
                MessageBox.Show("Попробуйте ввести число", "Ошибка!");
            }
        }


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
                    students.Add(new StudentDataClass() { 
                        ID = sID, FullName = line, Coefficient = 1.0,
                        attendAtLectures = 0, attendAtPractice = 0,
                        tasksCompleted = 0, questionsAnswered = 0
                    });
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
            foreach (string line in lines)
            {
                sID++;
                if (line != "")
                    questions.Add(new Questions() { ID = sID, QuestionText = line });
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
                AnswersDataTab.Visibility = Visibility.Hidden;
                studentsDict.Clear();
                MessageBox.Show("Добавьте вопросы или загрузите другой список.", "Вопросов меньше чем студентов!");
            }
            else
            {
                if (isQuestionListLoaded && isStudentListLoaded)
                {
                    RandomCreator.Visibility = Visibility.Visible;
                    AnswersDataTab.Visibility = Visibility.Visible;
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
                        MessageBox.Show("Место для сохранения не доступно.", "Ошибка");
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
        private void AnswersDataTab_RequestBringIntoView(object sender, RequestBringIntoViewEventArgs e)
        {
            StudentTab.IsEnabled = false;
            QuestionTab.IsEnabled = false;
            SpecialStudentsCorrection();
        }
        private void SpecialStudentsCorrection()
        {
            if (students == null)
                return;

            for (int i = 0; i < students.Count; i++)
            {
                foreach(var sStudent in specialStudents)
                {
                    if (students[i].FullName != null)
                        if (students[i].FullName.Contains(sStudent.Key))
                            students[i].Coefficient = sStudent.Value;
                }
            }
        }

        Dictionary<string, double> specialStudents = new Dictionary<string, double>()
        {
            { "Ношин", 1.4 },
            { "Павлов", 1.4 },
            { "Белоусов", 1.4 }
        };

        private void btnView_Click(object sender, RoutedEventArgs e)
        {
            StudentDataClass student = (StudentDataClass)((System.Windows.Controls.Button)e.Source).DataContext;
            if (
                student.attendAtLectures == 0 &&
                student.attendAtPractice == 0 &&
                student.tasksCompleted == 0 &&
                student.questionsAnswered == 0
            ) 
            {
                MessageBox.Show("Некорректные данные, либо студент не допускается до экзамена по умолчанию!");
                return;
            }
            for (int i = 0; i < students.Count; i++)
            {
                if (students[i].FullName != student.FullName)
                    continue;

                int tempMark = roundedMark(
                    student.Coefficient,
                    student.attendAtLectures,
                    student.attendAtPractice,
                    student.tasksCompleted,
                    student.questionsAnswered
                );
                students[i].mark = tempMark;
                MarksGrid.IsReadOnly = true;
                MarksGrid.Items.Refresh();
                MarksGrid.IsReadOnly = false;
            }
        }
        private int roundedMark(double Coefficient, int attendAtLectures, int attendAtPractice, int tasksCompleted, int questionsAnswered)
        {
            double grossMark = calculateMark(
                    Coefficient,
                    attendAtLectures,
                    attendAtPractice,
                    tasksCompleted,
                    questionsAnswered);
            if ( grossMark > 100 )
            {
                return 100;
            }
            return ((int)Math.Ceiling(grossMark));
            
        }
        private double calculateMark(double Coefficient, int attendAtLectures, int attendAtPractice, int tasksCompleted, int questionsAnswered)
        {
            if (
                attendAtLectures <= maxAttendanceLecture &&
                attendAtPractice <=maxAttendancePractice &&
                tasksCompleted <= maxAttendancePractice && 
                questionsAnswered <= 3
                )
            {
                if (attendAtLectures < 0 ||
                    attendAtPractice < 0 ||
                    tasksCompleted < 0 ||
                    questionsAnswered < 0
                    )
                {
                    MessageBox.Show("Введенные данные не совпадают. Оценка не была посчитана.", "Ошибка!");
                    return 0.0;
                } else
                {
                    return (
                        (Convert.ToDouble(attendAtLectures) / Convert.ToDouble(maxAttendanceLecture) * 25.0 +
                        Convert.ToDouble(attendAtPractice) / Convert.ToDouble(maxAttendancePractice) * 25.0 +
                        Convert.ToDouble(tasksCompleted) / Convert.ToDouble(maxAttendancePractice) * 30.0 +
                        Convert.ToDouble(questionsAnswered) / 3.0 * 20.0)
                        * Convert.ToDouble(Coefficient)
                    );
                }
            }
            MessageBox.Show("Введенные данные не совпадают. Оценка не была посчитана.", "Ошибка!");
            return 0.0;
        }
    }
}

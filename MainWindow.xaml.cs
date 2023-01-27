using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using Microsoft.Win32;
//using Xceed.Words.NET;
//using Xceed.Document.NET;
//using System.Data;
//using Utils.FileTools.cs

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
            string? fileName = FileTools.returnFileName();
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
            CheckIfEverythingLoaded();
        }
        private void CalculateMarks(object sender, RoutedEventArgs e)
        {
            FileTools.CreateExcelFile(students);
        }
        private void QuestionLoad(object sender, RoutedEventArgs e)
        {
            string? fileName = FileTools.returnFileName();

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
            CheckIfEverythingLoaded();
        }
        private void CheckIfEverythingLoaded()
        {
            if (questions.Count <= 3 && isQuestionListLoaded)
            {
                RandomCreator.Visibility = Visibility.Hidden;
                AnswersDataTab.Visibility = Visibility.Hidden;
                MessageBox.Show("Добавьте вопросы или загрузите другой список.", "Вопросов меньше чем студентов!");
                return;
            }

            if (students.Count < 1 && isStudentListLoaded)
            {
                RandomCreator.Visibility = Visibility.Hidden;
                AnswersDataTab.Visibility = Visibility.Hidden;
                MessageBox.Show("Добавьте хотя бы одного студента!", "Ошибка!");
                return;
            }

            if (isQuestionListLoaded && isStudentListLoaded)
            {
                RandomCreator.Visibility = Visibility.Visible;
                AnswersDataTab.Visibility = Visibility.Visible;
            }
        }
        private void CreateDoc(object sender, RoutedEventArgs e)
        {
            FileTools.CreateDocx(questions, students);
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
                        {
                            students[i].Coefficient = sStudent.Value;
                        }
                }
            }
        }

        public Dictionary<string, double> specialStudents = new()
        {
            { "Ношин ", 1.4 },
            { "Павлов ", 1.4 },
            { "Сигаев ", 1.4 }
        };

        private void BtnView_Click(object sender, RoutedEventArgs e)
        {
            StudentDataClass student;
            try
            {
                student = (StudentDataClass)((System.Windows.Controls.Button)e.Source).DataContext;
            }
            catch (Exception E)
            {
                return;
            }
            //if (
            //    (student.attendAtLectures == 0 &&
            //    student.attendAtPractice == 0 &&
            //    student.tasksCompleted == 0 &&
            //    student.questionsAnswered == 0) ||
            //    student.attendAtLectures > maxAttendanceLecture || 
            //    student.attendAtPractice > maxAttendancePractice || 
            //    student.tasksCompleted > maxAttendancePractice || 
            //    student.questionsAnswered > maxAttendancePractice
            //) 
            //{
            //    MessageBox.Show(
            //        $"Либо данные не введены либо завышены от изначальных значений.\n" +
            //        $"Максимальная посещаемость по лекциям: {maxAttendanceLecture}\n" +
            //        $"Максимальная посещаемость по практикам: {maxAttendancePractice}"
            //    );
            //    return;
            //}
            if (student.attendAtLectures > maxAttendanceLecture ||
            student.attendAtPractice > maxAttendancePractice ||
            student.tasksCompleted > maxAttendancePractice ||
            student.questionsAnswered > maxAttendancePractice
            ) 
            {
                MessageBox.Show(
                    $"Либо данные не введены либо завышены от изначальных значений.\n" +
                    $"Максимальная посещаемость по лекциям: {maxAttendanceLecture}\n" +
                    $"Максимальная посещаемость по практикам: {maxAttendancePractice}"
                );
                return;
            }

            int studentIndex = students.IndexOf(student);
            if (student.isIgnoreActive)
            {
                students[studentIndex].mark = 92;
            } else
            {
                students[studentIndex].mark = RoundedMark(
                    student.Coefficient,
                    student.attendAtLectures,
                    student.attendAtPractice,
                    student.tasksCompleted,
                    student.questionsAnswered
                );
            }
            students[studentIndex].isMarkSet = true;
            MarksGrid.IsReadOnly = true;
            MarksGrid.Items.Refresh();
            MarksGrid.IsReadOnly = false;
            //for (int i = 0; i < students.Count; i++)
            //{
            //    if (students[i].FullName != student.FullName)
            //        continue;


            //    int tempMark = roundedMark(
            //        student.Coefficient,
            //        student.attendAtLectures,
            //        student.attendAtPractice,
            //        student.tasksCompleted,
            //        student.questionsAnswered
            //    );
            //    students[i].mark = tempMark;
            //    students[i].isMarkSet = true;
            //    MarksGrid.IsReadOnly = true;
            //    MarksGrid.Items.Refresh();
            //    MarksGrid.IsReadOnly = false;
            //}
        }
        private int RoundedMark(double Coefficient, int attendAtLectures, int attendAtPractice, int tasksCompleted, int questionsAnswered)
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

        private void MarksGrid_CellEditEnding(object sender, System.Windows.Controls.DataGridCellEditEndingEventArgs e)
        {
            if (((StudentDataClass)e.Row.Item).isCoefActive)
            {
                ((StudentDataClass)e.Row.Item).Coefficient = 1.4;
            } else
            {
                ((StudentDataClass)e.Row.Item).Coefficient = 1.0;
            }
        }

        private void PrintMarks(object sender, RoutedEventArgs e)
        {
            FileTools.CreateExcelFile(students);
        }
    }
}
// TODO: Добавить рандомизацию вопросов
// TODO: Заставить некоторые вопросы отдавать определённым студентам
// TODO: Вывод всех оценок в DocX?

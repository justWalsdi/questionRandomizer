namespace QuestionRandomizer
{
    class StudentDataClass
    {
        public int ID { get; set; }
        public string? FullName { get; set; }
        public double Coefficient { get; set; }
        public int attendAtLectures { get; set; }
        public int attendAtPractice { get; set; }
        public int tasksCompleted { get; set; }
        public int questionsAnswered {get; set;}
        public int mark { get; set; }
        public bool isCoefActive { get; set; } = false;
        public bool isIgnoreActive { get; set; } = false;
        public bool isMarkSet { get; set; } = false;
    }
}

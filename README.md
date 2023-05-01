# FluentExcel

#Example Usage

    public class Program
    {
        static void Main(string[] args)
        {
            var items = new List<SampleClass>
            {
                new SampleClass { Id = 1, DateAdded = DateTime.Now, FirstName = "Mathew" },
                new SampleClass { Id = 2, DateAdded = DateTime.Now, FirstName = "John" },
                new SampleClass { Id = 3, DateAdded = DateTime.Now, FirstName = "Luke" },
                new SampleClass { Id = 4, DateAdded = DateTime.Now, FirstName = "Acts" },
                new SampleClass { Id = 5, DateAdded = DateTime.Now, FirstName = "Mark" },
            };

            ExcelBuilder
                .Begin()
                .AddWorkbook()
                //items to add and columns / column names used
                .AddWorkSheet(items.OrderBy(x => x.FirstName), item => item.Id, item => item.FirstName, item => item.DateAdded)
                .SaveWorkbook("example.xlsx")
                .End();
        }
    }

    public class SampleClass
    {
        public int Id { get; set; }
        public string FirstName { get; set; }
        public DateTime DateAdded { get; set; }
    }
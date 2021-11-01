# A simple tool that converts a POCO object to an excel file.



## Here's an example usage, where we assume that we have an excel file that contains person objects on each row

    

    class Program
    {
        static void Main(string[] args)
        {   
            var serializer = new ExcelSerializer();
            var data = serializer.ExcelFileToData<ExcelObj>(@"[write the directory of your excel file.]",1,100);

            foreach (var item in data)
            {
                Console.WriteLine($"Name = {item.Name},Age = {item.Age}");
            }
        }
    }

    public class Person
    {
        public dynamic Name { get; set; }
        public dynamic Age { get; set; }
    }

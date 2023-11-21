using OfficeOpenXml;

namespace ExcelAdaptation
{
    class Program
    {
        static void Main(string[] args)
        {
            DataRepository repository = new DataRepository("Data Source=(localdb)\\MSSQLLocalDB;Initial Catalog=NameOfDataBase;" +
                                                            "Integrated Security=True;Connect Timeout=30;Encrypt=False;");

            string filePath = "Path\\To\\File.xlsx";

            try
            {
                using (var stream = new MemoryStream())
                {
                    //Opening the file
                    using (var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                    {
                        fileStream.CopyTo(stream);
                    }

                    //Define working Excel Worksheet
                    using var package = new ExcelPackage(stream);
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                    //Counting rows
                    int rowCount = worksheet.Dimension.Rows;

                    //For each row, add the value in third column in the database
                    for (int row = 2; row <= rowCount; row++)
                    {
                        string Name = worksheet.Cells[row, 3].Value?.ToString();
                        TestObject rowModel = new TestObject(Name);
                        repository.AddTestObject(rowModel);
                        Console.WriteLine(rowModel.ToString() + "Ajouté avec succés");
                    }
                }
            }
            catch (Exception er)
            {
                Console.WriteLine("Une erreur s'est produite lors de l'importation :" + er.Message);
            }
        }
    }
}

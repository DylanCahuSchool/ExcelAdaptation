using OfficeOpenXml;

namespace ExcelAdaptation
{
    class Program
    {
        static void Main(string[] args)
        {
            // Création d'une instance de DataRepository pour accéder à la base de données
            DataRepository repository = new DataRepository("Data Source=(localdb)\\MSSQLLocalDB;Initial Catalog=XunitTest;" +
                                                            "Integrated Security=True;Connect Timeout=30;Encrypt=False;");

            // Chemin du fichier Excel à importer
            string filePath = "test\\test.xlsx";

            try
            {
                // Ouverture du fichier Excel en mode lecture
                using (var stream = new MemoryStream())
                {
                    using (var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                    {
                        // Copie du contenu du fichier Excel dans un MemoryStream
                        fileStream.CopyTo(stream);
                    }

                    // Définition de la feuille de calcul Excel à utiliser
                    using var package = new ExcelPackage(stream);
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                    // Comptage du nombre de lignes dans la feuille de calcul
                    int rowCount = worksheet.Dimension.Rows;

                    // Pour chaque ligne, ajout de la valeur de la troisième colonne dans la base de données
                    for (int row = 2; row <= rowCount; row++)
                    {
                        string Name = worksheet.Cells[row, 3].Value?.ToString();
                        TestObject rowModel = new TestObject(Name);
                        repository.AddTestObject(rowModel);
                        Console.WriteLine(rowModel.ToString() + " Ajouté avec succés");
                    }
                }
            }
            catch (Exception er)
            {
                // Affichage d'un message d'erreur en cas d'échec de l'importation
                Console.WriteLine("Une erreur s'est produite lors de l'importation :" + er.Message);
            }
        }
    }
}
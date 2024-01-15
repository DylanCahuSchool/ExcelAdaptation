using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;

namespace ExcelAdaptation
{
    public class DataRepository
    {
        private readonly string _connectionString; // Chaîne de connexion à la base de données

        public DataRepository(string connectionString)
        {
            _connectionString = connectionString; // Initialisation de la chaîne de connexion
        }

        public void AddTestObject(TestObject testObject)
        {
            using (SqlConnection connection = new SqlConnection(_connectionString))
            {
                connection.Open(); // Ouverture de la connexion à la base de données
                string sqlQuery = "INSERT INTO TestObject (TEST_NAME) VALUES (@Name)";

                using (SqlCommand command = new SqlCommand(sqlQuery, connection))
                {
                    command.Parameters.AddWithValue("@Name", testObject.getNom()); // Ajout du nom de l'objet en tant que paramètre

                    command.ExecuteNonQuery(); // Exécution de la requête SQL
                }
            }
        }
    }
}
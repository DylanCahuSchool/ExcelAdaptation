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
        private readonly string _connectionString;

        public DataRepository(string connectionString)
        {
            _connectionString = connectionString;
        }

        public void AddTestObject(TestObject testObject)
        {
            using (SqlConnection connection = new SqlConnection(_connectionString))
            {
                connection.Open();
                //INSERT INTO NameOfTable (FieldName) VALUES (@Name)
                string sqlQuery = "INSERT INTO TestObject (TEST_NAME) VALUES (@Name)";

                using (SqlCommand command = new SqlCommand(sqlQuery, connection))
                {
                    command.Parameters.AddWithValue("@Name", testObject.getNom());

                    command.ExecuteNonQuery();
                }
            }
        }
    }
}
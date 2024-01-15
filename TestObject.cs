using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelAdaptation
{
    public class TestObject
    {
        private int ID; // Identifiant de l'objet (non utilisé dans ce code)
        private string Name; // Nom de l'objet

        public TestObject(string Name)
        {
            this.Name = Name; // Initialisation du nom de l'objet
        }

        public string getNom()
        {
            return this.Name; // Récupération du nom de l'objet
        }
    }
}

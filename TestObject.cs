using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelAdaptation
{
    public class TestObject
    {
        private int ID;
        private string Name;

        public TestObject(string Name)
        {
            this.Name = Name;
        }

        public string getNom()
        {
            return this.Name;
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WinDev.Common.DataObjects
{
    public class Files
    {
        public int column;
        public string name;

        public Files(string name, int column)
        {
            this.column = column;
            this.name = name;
        }

        public override string ToString()
        {
            return new String('|', column) + name;
        }
    }
}

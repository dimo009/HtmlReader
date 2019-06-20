using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace urlReader.Classes
{
    public class Package
    {
        public  string Name { get; set; }

        public List<string> Versions { get; set; }

        public List<string> Components { get; set; }
    }
}

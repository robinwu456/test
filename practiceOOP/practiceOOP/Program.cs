using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace practiceOOP {
    class Program {
        static void Main(string[] args) {

            area a = new area(6, 8);

            int aa = a.AreaCount();

            Console.WriteLine(aa);

            Console.ReadKey();
        }
    }
}

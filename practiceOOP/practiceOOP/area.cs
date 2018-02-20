using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace practiceOOP {
    class area {

        private int length;
        private int width;

        public area(int length, int width) {

            this.length = length;
            this.width = width;
        }

        public int AreaCount() {
            return length * width;
        }
    }
}

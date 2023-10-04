using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace rgr1
{
    internal class Students
    {
        private string name;
        private int mat;
        private int fiz;
        private int him;
        private int rus;
        private string work;

        public string Name { get; set; }
        public int Mat
        { set { mat = value; } }
        public int Fiz
        { set { fiz = value; } }
        public int Him
        { set { him = value; } }
        public int Rus
        { set { rus = value; } }
        public string Work
        {
            get { return work; }
            set { work = value; }
        }
        public Students(string name, int mat, int fiz, int him, int rus, string work)
        {
            this.Name = name;
            this.Mat = mat;
            this.Fiz = fiz;
            this.Him = him;
            this.Rus = rus;
            this.Work = work;
        }
        public Students() { }
        public double Stipa()
        {
            double st = 0;
            double baza = 1000;
            int[] a = new int[4];
            a[0] = mat;
            a[1] = fiz;
            a[2] = him;
            a[3] = rus;
            string wrk1 = "староста";
            string wrk2 = "студсовет";
            int s2 = 0;
            int s3 = 0;
            int s4 = 0;
            int s5 = 0;
            for (int i = 0; i < 4; i++)
            {
                if (a[i] == 2)
                    s2++;
                if (a[i] == 3)
                    s3++;
                if (a[i] == 4)
                    s4++;
                if (a[i] == 5)
                    s5++;
            }
            if (s2 == 0 && (s4 >= 1 || s5 >= 1))
                st = baza;
            if (s2 == 0 && s3 == 0 && s5 >= 1 && s4 >= 1)
                st = baza * 1.2;
            if (s5 == 4)
                st = baza * 1.3;
            if (work == wrk1 && s2 == 0)
                st = st * 1.2;
            if (work == wrk2 && s2 == 0 && s3 == 0)
                st = st * 1.1;
            return st;
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CNR
{
    internal class HGAdminC
    {
        int u, d, t, c, ci, subu, subd, subt, subc, subci;
        double un, dos, tr, cu, cin, total;

        public int U { get => u; set => u = value; }
        public int D { get => d; set => d = value; }
        public int T { get => t; set => t = value; }
        public int C { get => c; set => c = value; }
        public int Ci { get => ci; set => ci = value; }
        public int Subu { get => subu; set => subu = value; }
        public int Subd { get => subd; set => subd = value; }
        public int Subt { get => subt; set => subt = value; }
        public int Subc { get => subc; set => subc = value; }
        public int Subci { get => subci; set => subci = value; }
        public double Total { get => total; set => total = value; }
        public double Un { get => un; set => un = value; }
        public double Dos { get => dos; set => dos = value; }
        public double Tr { get => tr; set => tr = value; }
        public double Cu { get => cu; set => cu = value; }
        public double Cin { get => cin; set => cin = value; }

        public void totales()
        {
            Subu = U * 1;
            Subd = D * 2;
            Subt = T * 3;
            Subc = C * 4;
            Subci = Ci * 5;
            Total = (un + dos + tr + cu + cin) / 5;
        }
        
    }
}
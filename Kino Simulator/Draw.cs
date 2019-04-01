using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kino_Simulator
{
   
    public class Draw
    {
        public string drawTime { get; set; }
        public int drawNo { get; set; }
        public List<int> results { get; set; }
    }

    public class Draws
    {
        public List<Draw> draw { get; set; }
    }

    public class RootObject
    {
        public Draws draws { get; set; }
        public Draw draw { get; set; }

    }
}

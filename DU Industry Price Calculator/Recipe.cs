using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DU_Industry_Price_Calculator
{
    public class Recipe
    {
        public string Name { get; set; }
        public int Tier { get; set; }
        public string Type { get; set; }
        public double Mass { get; set; }
        public double Volume { get; set; }
        public double OutputQuantity { get; set; }
        public double Time { get; set; }
        public Dictionary<string, double> Byproducts { get; set; } = new Dictionary<string, double>();
        public List<string> Industries { get; set; } = new List<string>();
        public Dictionary<string, double> Input { get; set; } = new Dictionary<string, double>();
        public double Price { get; set; } // Buy and Sell prices are assumedly pretty much the same
    }
}

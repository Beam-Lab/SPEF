using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BeamLab.SPEF.Test.Models;

namespace BeamLab.SPEF.Test
{
    public class TestContext : SPEFContext
    {
        public TestContext(string contextUrl, string mainVariationLabel) : base(contextUrl, mainVariationLabel)
        {
        }

        public List<Category> Categories { get; set; }
        public List<News> News { get; set; }
    }
}

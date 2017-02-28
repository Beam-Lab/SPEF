using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BeamLab.SPEF.Test.Models
{
    [SPEFList(Title = "News", Description = "News del Portale")]
    public class News : SPEFListItem
    {
        [SPEFNumericField]
        public int LikeCount { get; set; }

        [SPEFTextField(MaxLength = 150)]
        public string Text { get; set; }

        [SPEFLookupField]
        public Category Category { get; set; }
    }
}

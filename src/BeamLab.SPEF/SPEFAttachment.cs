using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BeamLab.SPEF
{
    public class SPEFAttachment
    {
        public byte[] FileContent { get; set; }
        public string FileName { get; set; }
        public string ServerRelativeUrl { get; set; }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BeamLab.SPEF.Models;
using System.Runtime.Serialization;

namespace BeamLab.SPEF
{
    public class SPEFDocListItem : SPEFListItem
    {
        public byte[] FileContent { get; set; }
        public bool FileOverwrite { get; set; }
        public string Name { get; set; }

        #region internal

        [SPEFTextField(Name = "File_x0020_Type", Title = "File_x0020_Type", Ignore = true, Readonly = true)]
        public string FileType { get; set; }
        [SPEFTextField(Name = "FileLeafRef", Title = "FileLeafRef", Ignore = true, Readonly = true)]
        public string FileName { get; set; }
        [SPEFTextField(Name = "FileRef", Title = "FileRef", Ignore = true, Readonly = true)]
        public string FileUrl { get; set; }
        [SPEFTextField(Name = "FileDirRef", Title = "FileDirRef", Ignore = true, Readonly = true)]
        public string DirectoryUrl { get; set; }
        [SPEFNumericField(Name = "File_x0020_Size", Title = "File_x0020_Size", Ignore = true, Readonly = true)]
        public int FileSizeDisplay { get; set; }

        #endregion
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BeamLab.SPEF.Models;
using BeamLab.SPEF.Extensions;
    
namespace BeamLab.SPEF
{
    public class SPEFItem
    {
        [SPEFNumericField(Ignore = true, Readonly = true)]
        public int ID { get; set; }

        public string DisplayField { get; set; }
    }

    public class SPEFListItem : SPEFItem
    {
        public SPEFListItem()
        {
            Attachments = new List<SPEFAttachment>();
        }

        [SPEFTextField(Ignore = true)]
        public string Title { get; set; }
        [SPEFDateTimeField(Ignore = true)]
        public DateTime Modified { get; set; }
        [SPEFDateTimeField(Ignore = true, Readonly = true)]
        public DateTime Created { get; set; }

        #region Author
        public void SetAuthor(SPEFUser author)
        {
            Author = author;
            AuthorUpdated = true;
        }
        internal bool AuthorUpdated = false;
        [SPEFUserField(Ignore = true)]
        public SPEFUser Author { get; internal set; }
        #endregion

        #region Editor
        public void SetEditor(SPEFUser editor)
        {
            Editor = editor;
            EditorUpdated = true;
        }
        internal bool EditorUpdated = false;
        [SPEFUserField(Ignore = true)]
        public SPEFUser Editor { get; internal set; }
        #endregion

        public List<SPEFAttachment> Attachments { get; set; } 
    }
}

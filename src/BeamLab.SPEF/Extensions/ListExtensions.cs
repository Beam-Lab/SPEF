using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BeamLab.SPEF.Extensions
{
    public static class ListExtensions
    {
        public static ListItem AddItem(this List list)
        {
            var itemCreationInformation = new ListItemCreationInformation();
            var oListItem = list.AddItem(itemCreationInformation);

            return oListItem;
        }
    }
}

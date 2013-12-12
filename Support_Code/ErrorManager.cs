using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Windows.Forms;

namespace ExcelEditor
{
    public partial class ErrorManager
    {
        internal static void ErrorNotification(string errorKey, object obj)
        {
            switch (errorKey)
            {
                case "NULL" :
                    MessageBox.Show("Null pointer for object " + obj.ToString());
                    break;
                case "GENERAL" :
                    MessageBox.Show(obj.ToString());
                    break;    
                default: break;
            }

        }
    }
}

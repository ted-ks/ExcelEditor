using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace ExcelEditor
{
    public partial class FileSelector
    {
        internal static bool ValidateOpenedFile(string FileName)
        {
            string extension = Path.GetExtension(FileName);
            if (extension == ".xlsx") return true;
            return false;
        }

    }
}

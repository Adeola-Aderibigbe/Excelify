﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace Excelify.Services.Extensions
{
    public static class ExcelifyExtension
    {
        public static void ToFile(this byte[] workSheet, string path)
        {
            workSheet.WriteToFile(path);
        }

    }
}

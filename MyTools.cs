using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kwerendy
{
    public static class MyTools
    {
        public static List<Entry> allEntries { get; set; } = new List<Entry>();

        public static List<Entry> registrations { get; set; } = new List<Entry>();

        public static string signatureFilePath = "";
    }
}

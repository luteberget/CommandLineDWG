using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace dwginfo
{
    static class StringTools
    {
        public static string FormatCountAndExamples(ICollection<string> names, int width)
        {
            int count = names.Count;
            string retval = String.Format("{0}", count);
            if (names.Count > 0)
            {
                retval += " (";
                string end = ")";
                retval += FormatExamples(names, width - retval.Length - end.Length);
                retval += end;
            }

            return retval;
        }

        public static string FormatExamples(ICollection<string> names, int width)
        {
            int remaining = width;
            if (remaining < 3) { return ""; }
            string end = "...";
            string separator = ", ";
            string retval = "";

            for (int i = 0; i < names.Count; i++)
            {
                int required = names.ElementAt(i).Length;
                if (i + 1 < names.Count)
                {
                    required += end.Length + separator.Length;
                }
                if (remaining >= required)
                {
                    retval += names.ElementAt(i);
                    remaining -= names.ElementAt(i).Length;
                    if (i + 1 < names.Count)
                    {
                        retval += separator;
                        remaining -= separator.Length;
                    }
                }
                else
                {
                    retval += end;
                    break;
                }
            }

            return retval;

        }
    }
}

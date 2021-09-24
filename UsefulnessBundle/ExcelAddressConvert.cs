using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UsefulnessBundle
{
    class ExcelAddressConvert
    {
        /// <summary>
        /// 将指定的自然数转换为26进制表示。映射关系：[1-26] ->[A-Z]。
        /// </summary>
        /// <param name="n">自然数（如果无效，则返回空字符串）。</param>
        /// <returns>26进制表示。</returns>
        public static string ToNumberSystem26(int n)
        {
            int temp = n;
            string s = string.Empty;
            while (temp > 0)
            {
                int m = temp % 26;
                if (m == 0) m = 26;
                s = (char)(m + 64) + s;
                temp = (temp - m) / 26;
            }
            return s;
        }

        /// <summary>
        /// 将指定的自然数转换为26进制表示。映射关系：[1-26] ->[A-Z]。
        /// </summary>
        /// <param name="n">自然数（如果无效，则返回空字符串）。</param>
        /// <param name="startWithZero">数字起点是不是从0开始，可能会改变映射的范围</param>
        /// <returns>26进制表示。</returns>
        public static string ToNumberSystem26(int n, bool startWithZero)
        {
            int temp;
            string s = string.Empty;
            if (startWithZero)
            {
                temp = n + 1;
            }
            else
            {
                temp = n;
            }
            while (temp > 0)
            {
                int m = temp % 26;
                if (m == 0) m = 26;
                s = (char)(m + 64) + s;
                temp = (temp - m) / 26;
            }
            return s;
        }

        /// <summary>
        /// 将指定的26进制表示转换为自然数。映射关系：[A-Z] ->[1-26]。
        /// </summary>
        /// <param name="s">26进制表示（如果无效，则返回0）。</param>
        /// <returns>自然数。</returns>
        public static int FromNumberSystem26(string s)
        {
            if (string.IsNullOrEmpty(s)) return 0;
            int n = 0;
            for (int i = s.Length - 1, j = 1; i >= 0; i--, j *= 26)
            {
                char c = Char.ToUpper(s[i]);
                if (c < 'A' || c > 'Z') return 0;
                n += ((int)c - 64) * j;
            }
            return n;
        }

        /// <summary>
        /// 将指定的26进制表示转换为自然数。映射关系：[A-Z] ->[1-26]。
        /// </summary>
        /// <param name="s">26进制表示（如果无效，则返回0）。</param>
        /// <param name="startWithZero">数字起点是不是从0开始，可能会改变映射的范围</param>
        /// <returns>自然数。</returns>
        public static int FromNumberSystem26(string s, bool startWithZero)
        {
            if (string.IsNullOrEmpty(s)) return 0;
            int n = 0;
            for (int i = s.Length - 1, j = 1; i >= 0; i--, j *= 26)
            {
                char c = Char.ToUpper(s[i]);
                if (c < 'A' || c > 'Z') return 0;
                n += ((int)c - 64) * j;
            }

            if (startWithZero)
            {
                return n - 1;
            }
            else
            {
                return 0;
            }
        }
    }
}

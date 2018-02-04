using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeBrute
{
    public class BruteObj
    {
        private string m_filePath = "";
        private string m_wordlistPath = "";
        private UInt64 m_start = 0;
        private UInt64 m_end = 0;

        public BruteObj(string filePath, string wordlistPath, UInt64 start, UInt64 end)
        {
            this.m_filePath = filePath;
            this.m_wordlistPath = wordlistPath;
            this.m_start = start;
            this.m_end = end;
        }
        public string getWordlistPath()
        {
            return this.m_wordlistPath;
        }
        public string getPath()
        {
            return this.m_filePath;
        }
        public UInt64 getStart()
        {
            return this.m_start;
        }
        public UInt64 getEnd()
        {
            return this.m_end;
        }
    }
}

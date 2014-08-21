using Microsoft.Build.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EspritLogger
{
    public class ProjectData
    {
        private int m_count;
        private BuildWarningEventArgs m_warningArgs;

        public ProjectData() 
        {
            m_count = 0;
        }

        public ProjectData(BuildWarningEventArgs warningArgs) 
        {
            m_warningArgs = warningArgs;
            m_count = 1;
        }

        public void AddWarningArgs(BuildWarningEventArgs warningArgs) 
        {
            m_warningArgs = warningArgs;
            m_count++;
        }

        public BuildWarningEventArgs GetWarningArgs()
        {
            return m_warningArgs;
        }

        public int GetWarningCount()
        {
            return m_count;
        }
    }
}

using Microsoft.Build.Framework;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace EspritLogger
{
    public class CustomTextLogger
    {

        private StreamWriter streamWriter;

        public CustomTextLogger(string logTxt)
        {
            try
            {
                // Open the file 
                this.streamWriter = new StreamWriter(logTxt);
                streamWriter.WriteLine("=============Code Analysis Warnings=============");
            }
            catch (Exception ex)
            {
                if
                (
                    ex is UnauthorizedAccessException
                    || ex is ArgumentNullException
                    || ex is PathTooLongException
                    || ex is DirectoryNotFoundException
                    || ex is NotSupportedException
                    || ex is ArgumentException
                    || ex is SecurityException
                    || ex is IOException
                )
                {
                    throw new LoggerException("Failed to create log file: " + ex.Message);
                }
                else
                {
                    // Unexpected failure 
                    throw;
                }
            }
        }

        public void WriteLineWithSenderAndMessage(string line, BuildEventArgs e)
        {      
                WriteLine(e.SenderName + " : " + line, e);
        }

        /// <summary> 
        /// Just write a line to the log 
        /// </summary> 
        public void WriteLine(string line, BuildEventArgs e)
        {
            streamWriter.WriteLine(line + e.Message);
        }

        public void Close() 
        {
            streamWriter.Close();
        }

    }
}

using System;
using System.IO;
using System.Security;
using Microsoft.Build.Framework;
using Microsoft.Build.Utilities;
using System.Collections.Generic;

namespace EspritLogger
{

	public class BasicFileLogger : Logger
	{

        private Dictionary<string, ProjectData> m_projectWarnings;
        private string m_logXls;
        private string m_logTxt;

        private CustomTextLogger m_textLogger;
        private bool m_buildFailed;

		public override void Initialize(IEventSource eventSource)
		{

            m_buildFailed = false;

			if (null == Parameters)
			{
				throw new LoggerException("Log files were not set.");
			}
			string[] parameters = Parameters.Split(';');

            //if (parameters.Length > 0) 
            //{
            //    m_logXls = parameters[0];
            //    if (parameters.Length > 1) 
            //    {
            //        m_logTxt = parameters[1];
            //    }
            //}

            m_logXls = parameters[0];
            m_logTxt = parameters[1];
            if (String.IsNullOrEmpty(m_logXls))
			{
				throw new LoggerException("Xls log file was not set.");
			}

            if (String.IsNullOrEmpty(m_logTxt))
            {
                throw new LoggerException("Txt log file was not set.");
            }

			if (parameters.Length > 3)
			{
				throw new LoggerException("Too many parameters passed.");
			}

            

            //Open excel log file
            m_projectWarnings = new Dictionary<String, ProjectData>();
            //Open txt log file 
            m_textLogger = new CustomTextLogger(m_logTxt);

			// For brevity, we'll only register for certain event types. Loggers can also
			// register to handle TargetStarted/Finished and other events.
            eventSource.ProjectStarted += new ProjectStartedEventHandler(eventSource_ProjectStarted);
            eventSource.TaskStarted += new TaskStartedEventHandler(eventSource_TaskStarted);
            eventSource.MessageRaised += new BuildMessageEventHandler(eventSource_MessageRaised);
            eventSource.WarningRaised += new BuildWarningEventHandler(eventSource_WarningRaised);
            eventSource.ErrorRaised += new BuildErrorEventHandler(eventSource_ErrorRaised);
            eventSource.ProjectFinished += new ProjectFinishedEventHandler(eventSource_ProjectFinished);
		}

		void eventSource_ErrorRaised(object sender, BuildErrorEventArgs e)
		{
			// BuildErrorEventArgs adds LineNumber, ColumnNumber, File, amongst other parameters
            m_buildFailed = true;
            //string line = String.Format(": ERROR {0}({1},{2}): ", e.File, e.LineNumber, e.ColumnNumber);
            //m_textLogger.WriteLineWithSenderAndMessage(line, e);
		}
		
		void eventSource_WarningRaised(object sender, BuildWarningEventArgs e)
		{
            if (String.Compare(e.SenderName, "Microsoft.Build.Tasks.CodeAnalysis", true) == 0)
            {
                string line = String.Format("{0} : Warning {1}({2},{3}): ", e.Code, e.File, e.LineNumber, e.ColumnNumber);
                m_textLogger.WriteLine(line, e);

                string projectName = Path.GetFileNameWithoutExtension(e.ProjectFile);
                //m_projectWarnings[projectName].AddWarningArgs(e);
                if (m_projectWarnings.ContainsKey(projectName))
                {
                    m_projectWarnings[projectName].AddWarningArgs(e);
                }
                else
                {
                    ProjectData data = new ProjectData(e);
                    m_projectWarnings.Add(projectName, data);
                }
            }
		}

        void eventSource_MessageRaised(object sender, BuildMessageEventArgs e)
        {
            // BuildMessageEventArgs adds Importance to BuildEventArgs 
            // Let's take account of the verbosity setting we've been passed in deciding whether to log the message 
            //if ((e.Importance == MessageImportance.High && IsVerbosityAtLeast(LoggerVerbosity.Minimal))
            //    || (e.Importance == MessageImportance.Normal && IsVerbosityAtLeast(LoggerVerbosity.Normal))
            //    || (e.Importance == MessageImportance.Low && IsVerbosityAtLeast(LoggerVerbosity.Detailed))
            //    )
            //{
            //    m_textLogger.WriteLineWithSenderAndMessage(String.Empty, e);
            //}
        }

        void eventSource_TaskStarted(object sender, TaskStartedEventArgs e)
        {
            // TaskStartedEventArgs adds ProjectFile, TaskFile, TaskName 
            // To keep this log clean, this logger will ignore these events.
        }

        void eventSource_ProjectStarted(object sender, ProjectStartedEventArgs e)
        {
            // ProjectStartedEventArgs adds ProjectFile, TargetNames 
            // Just the regular message string is good enough here, so just display that.
            //m_textLogger.WriteLine(String.Empty, e);
        }

        void eventSource_ProjectFinished(object sender, ProjectFinishedEventArgs e)
        {
            // The regular message string is good enough here too.
        }

        public override void Shutdown()
		{
            // Done logging, let go of the file
            m_textLogger.Close();
            //Update regression logs only if build sucseeded
            ExcelFileHandler excelFileHandler = new ExcelFileHandler(m_projectWarnings, m_logXls);
            //excelFileHandler.UpdateLog(!m_buildFailed);
            excelFileHandler.UpdateLog(true);
            excelFileHandler.SaveLogFileAs();
		}
	}
}

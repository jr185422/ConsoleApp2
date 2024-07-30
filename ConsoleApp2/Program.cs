using System;
using System.IO;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;

namespace CrystalReportsResave
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.Write("Enter the directory path for .rpt files: ");
            string reportDirectory = Console.ReadLine();

            Console.Write("Enter the directory path to save updated .rpt files: ");
            string newReportDirectory = Console.ReadLine();

            string logFilePath = Path.Combine(newReportDirectory, "report_versions_log.txt");
            string errorLogFilePath = Path.Combine(newReportDirectory, "report_errors_log.txt");
            string summaryLogFilePath = Path.Combine(newReportDirectory, "report_summary_log.txt");

            // Ensure the new report directory exists
            if (!Directory.Exists(newReportDirectory))
            {
                Directory.CreateDirectory(newReportDirectory);
            }

            Console.Write("Enter server name: ");
            string serverName = Console.ReadLine();
            Console.Write("Enter database name: ");
            string databaseName = Console.ReadLine();
            Console.Write("Enter user ID: ");
            string userId = Console.ReadLine();
            Console.Write("Enter password: ");
            string password = Console.ReadLine();

            ConnectionInfo connectionInfo = new ConnectionInfo
            {
                ServerName = serverName,
                DatabaseName = databaseName,
                UserID = userId,
                Password = password
            };

            int totalFiles = 0;
            int processedFiles = 0;
            int errorFiles = 0;

            ProcessReports(reportDirectory, newReportDirectory, logFilePath, errorLogFilePath, connectionInfo, ref totalFiles, ref processedFiles, ref errorFiles);

            string summary = $"Total .rpt files: {totalFiles}\n" +
                             $"Successfully processed files: {processedFiles}\n" +
                             $"Files with errors: {errorFiles}\n" +
                             $"Error log file: {errorLogFilePath}";

            Console.WriteLine(summary);
            LogSummary(summaryLogFilePath, summary);
        }

        static void ProcessReports(string reportDirectory, string newReportDirectory, string logFilePath, string errorLogFilePath, ConnectionInfo connectionInfo, ref int totalFiles, ref int processedFiles, ref int errorFiles)
        {
            foreach (var file in Directory.GetFiles(reportDirectory, "*.rpt", SearchOption.AllDirectories))
            {
                totalFiles++;
                if (OpenAndResaveRpt(file, newReportDirectory, logFilePath, errorLogFilePath, connectionInfo))
                {
                    processedFiles++;
                }
                else
                {
                    errorFiles++;
                }
            }
        }

        static bool OpenAndResaveRpt(string filePath, string newReportDirectory, string logFilePath, string errorLogFilePath, ConnectionInfo connectionInfo)
        {
            ReportDocument reportDocument = new ReportDocument();

            try
            {
                // Load the report
                reportDocument.Load(filePath);
                Console.WriteLine($"Report {filePath} loaded successfully.");

                // Set the database login credentials
                SetDBLogonForReport(connectionInfo, reportDocument);
            
                // Load data into the report (from preview)
                reportDocument.VerifyDatabase();
                reportDocument.Refresh();
                Console.WriteLine("Data loaded into the report.");

                // Enable "Save Data with Report"
                reportDocument.ReportOptions.EnableSaveDataWithReport = true;
                // Handle missing parameter values for the main report and subreports
                SetDefaultParameterValues(reportDocument);


                // Save the report with a new name in the new directory
                string newFileName = Path.GetFileNameWithoutExtension(filePath) + ".rpt";
                string newFilePath = Path.Combine(newReportDirectory, newFileName);

                reportDocument.SaveAs(newFilePath, true);
                Console.WriteLine($"Report saved as {newFilePath}.");

                // Log Crystal Reports version
                string crystalVersion = GetCrystalReportsVersion();
                LogVersion(logFilePath, filePath, crystalVersion);
                Console.WriteLine($"Report {filePath} is using Crystal Reports version: {crystalVersion}");

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file {filePath}: {ex.Message}");
                LogError(errorLogFilePath, filePath, ex.Message);
                return false;
            }
            finally
            {
                // Close the report
                reportDocument.Close();
                reportDocument.Dispose();
            }
        }

        static void SetDBLogonForReport(ConnectionInfo connectionInfo, ReportDocument reportDocument)
        {
            Tables tables = reportDocument.Database.Tables;

            foreach (Table table in tables)
            {
                TableLogOnInfo tableLogOnInfo = table.LogOnInfo;
                tableLogOnInfo.ConnectionInfo = connectionInfo;
                table.ApplyLogOnInfo(tableLogOnInfo);
            }

            foreach (ReportDocument subReport in reportDocument.Subreports)
            {
                Tables subReportTables = subReport.Database.Tables;
                foreach (Table table in subReportTables)
                {
                    TableLogOnInfo tableLogOnInfo = table.LogOnInfo;
                    tableLogOnInfo.ConnectionInfo = connectionInfo;
                    table.ApplyLogOnInfo(tableLogOnInfo);
                }
            }
        }

        static void SetDefaultParameterValues(ReportDocument reportDocument)
        {
            foreach (ParameterFieldDefinition param in reportDocument.DataDefinition.ParameterFields)
            {
                if (param.ReportName == "" || param.ReportName == reportDocument.Name) // This checks if it's a main report parameter
                {
                    ParameterValues currentValues = param.CurrentValues;

                    if (currentValues.Count == 0)
                    {
                        ParameterDiscreteValue defaultValue = new ParameterDiscreteValue();
                        defaultValue.Value = null;
                        currentValues.Add(defaultValue);
                        param.ApplyCurrentValues(currentValues);
                    }
                }
            }
        }

        static string GetCrystalReportsVersion()
        {
            // Retrieve the version of the CrystalDecisions.CrystalReports.Engine assembly
            var crystalReportsAssembly = typeof(ReportDocument).Assembly;
            var version = crystalReportsAssembly.GetName().Version;
            return version.ToString();
        }

        static void LogVersion(string logFilePath, string filePath, string version)
        {
            try
            {
                using (StreamWriter sw = new StreamWriter(logFilePath, true))
                {
                    sw.WriteLine($"Report {filePath} is using Crystal Reports version: {version}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to write to log file: {ex.Message}");
            }
        }

        static void LogError(string errorLogFilePath, string filePath, string errorMessage)
        {
            try
            {
                using (StreamWriter sw = new StreamWriter(errorLogFilePath, true))
                {
                    sw.WriteLine($"Error processing file {filePath}: {errorMessage}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to write to error log file: {ex.Message}");
            }
        }

        static void LogSummary(string summaryLogFilePath, string summary)
        {
            try
            {
                using (StreamWriter sw = new StreamWriter(summaryLogFilePath, true))
                {
                    sw.WriteLine(summary);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to write to summary log file: {ex.Message}");
            }
        }
    }
}
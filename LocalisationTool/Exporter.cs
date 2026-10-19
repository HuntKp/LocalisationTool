using LocalisationTool;
using Language;
using System;
using System.Text;
using System.Text.RegularExpressions;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ImportExport
{
    internal class Exporter
    {
        public MainWindow MainWindow { get; set; }

        public void ExportToExcel_Interop(string sFilePath, List<LocalisationProject> lProjects)
        {
			//delete any files which already exist with the same name
            if ((File.Exists(sFilePath)))
                File.Delete(sFilePath);

            var excelApplication = new Excel.Application();
            var excelWorkbook = excelApplication.Workbooks.Add();

            foreach (var project in lProjects)
            {
                var excelWorksheet = excelWorkbook.Sheets.Add();
                excelWorksheet.Name = project.ProjectName ?? "Project";

                //Gather language IDs
                var allIds = new HashSet<string>(project.Languages.SelectMany(lang => lang.LocalisationEntries.Select(e => e.ID)));

                // Prepare language headers
                var languageNames = project.Languages.Select(l => l.Language).ToList(); 

                // Write header row: ID, then one column per language
                excelWorksheet.Cells[1, 1] = "ID";
                for (int i = 0; i < languageNames.Count; i++)  
                {
                    excelWorksheet.Cells[1, i + 2] = languageNames[i];
                } 

                // Write data rows 
                int row = 2;
                foreach (var id in allIds)
                {
                    excelWorksheet.Cells[row, 1] = id;
                    for (int col = 0; col < languageNames.Count; col++)
                    {
                        var lang = project.Languages[col];
                        var entry = lang.LocalisationEntries.FirstOrDefault(e => e.ID == id);
                        excelWorksheet.Cells[row, col + 2] = entry?.Text ?? "";
                    }
                    row++;
                }

                // Format header row
                var rangeHeader = excelWorksheet.Range["A1", excelWorksheet.Cells[1, languageNames.Count + 1]];
                rangeHeader.Font.Bold = true;
                rangeHeader.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue);
                rangeHeader.Font.Size = 12;
                rangeHeader.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                // AutoFit columns
                ((Excel.Range)excelWorksheet.Columns).AutoFit();

                //Freeze top two rows and columns
                excelWorksheet.Application.ActiveWindow.SplitRow = 1;
                excelWorksheet.Application.ActiveWindow.SplitColumn = 1;
                excelWorksheet.Application.ActiveWindow.SplitColumn = 2;
                excelWorksheet.Application.ActiveWindow.FreezePanes = true;
                //Unlock all cells
                excelWorksheet.Cells.Locked = false;
                //Now we lock the cells we want to lock, using a range
                excelWorksheet.Range["A:A"].Locked = true; // Column 1, A
                //Unlock cells in a range
                excelWorksheet.Range["B:B"].Locked = false; // Column 2, B
                //Now lock the rows we want to lock
                excelWorksheet.Rows[1].Locked = true; //Row 1
                excelWorksheet.Protect(Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            }

            excelWorkbook.SaveAs(sFilePath);
            excelWorkbook.Close();
            excelApplication.Quit();
        }

        public void ExportToRCFiles(LocalisationContainer localisationContainer)
        {
            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = "Select a folder containing .rc files";
                dialog.ShowNewFolderButton = false;
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    string selectedPath = dialog.SelectedPath;

                    //our projects
                    List<LocalisationProject> projects = localisationContainer.Projects;

                    //find all rc files, then only loop through the onese references in our projects
                    List<string> rcFileList = FindRcFiles(selectedPath/*, allowedFolders*/);

                    if (rcFileList.Count == 0)
                    {
                        Console.WriteLine("No .rc files found in the specified folder.");
                        return;
                    }
                    else
                    {
                        foreach (string sPath in rcFileList)
                        {
                            string sFileNameNoExt = Path.GetFileNameWithoutExtension(sPath);

                            // Check for a matching project
                            var matchingProject = projects.FirstOrDefault(
                                p => string.Equals(Path.GetFileNameWithoutExtension(p.ProjectName), sFileNameNoExt, StringComparison.OrdinalIgnoreCase)
                            );

                            if (matchingProject != null)
                            {
                                // Use the selectedPath as needed, e.g.:
                                ExportToRCFile(matchingProject, sPath);
                            }
                        }
                    }
                }
            }
        }

        public void ExportToRCFile(LocalisationProject localisationProject, string selectedPath)
        {
            var lines = File.ReadAllLines(selectedPath);
            var outputLines = new List<string>();

            // Prepare a dictionary for quick language lookup
            var languageDict = localisationProject.Languages
                .ToDictionary(l => l.Language, l => l, StringComparer.OrdinalIgnoreCase);

            string currentLanguage = null;
            LocalisationLanguage currentLocalisationLanguage = null;
            var languageHeaderPattern = new Regex(@"^\s*LANGUAGE\s+(.+)$", RegexOptions.IgnoreCase);
            var entryPattern = new Regex(@"^\s*([A-Za-z0-9_]+)\s+""(.*)""\s*$");

            foreach (var line in lines)
            {
                var trimmed = line.Trim();

                // Detect LANGUAGE statement
                var langMatch = languageHeaderPattern.Match(trimmed);
                if (langMatch.Success)
                {
                    currentLanguage = langMatch.Groups[1].Value.Trim();
                    // Try to match using your language mapping logic if needed
                    languageDict.TryGetValue(MainWindow.LanguageStringMatch(currentLanguage), out currentLocalisationLanguage);
                    outputLines.Add(line);
                    continue;
                }

                // Detect ID "Text" entry
                var entryMatch = entryPattern.Match(line);
                if (entryMatch.Success && currentLocalisationLanguage != null
                    //temp illegal options
                    && !line.Equals("POPUP")
                    && !line.Equals("CAPTION")
                    && !line.Equals("BLOCK")
                    && !line.Equals("VALUE"))
                {
                    string id = entryMatch.Groups[1].Value;
                    var entry = currentLocalisationLanguage.LocalisationEntries
                        .FirstOrDefault(e => e.ID == id);

                    if (entry != null)
                    {
                        // Replace the text with the value from LocalisationEntry
                        outputLines.Add($"{id} \"{entry.Text}\"");
                    }
                    else
                    {
                        // No match, keep the original line
                        outputLines.Add(line);
                    }
                }
                else
                {
                    // Not a language or entry line, keep as is
                    outputLines.Add(line);
                }
            }

            File.WriteAllLines(selectedPath, outputLines, Encoding.UTF8);
        }

        public List<string> FindRcFiles(string folderPath/*, List<string> allowedFolders*/)
        {
            if (!Directory.Exists(folderPath))
                return new List<string>();

            var rcFiles = Directory.GetFiles(folderPath, "*.rc", SearchOption.AllDirectories);
            var filteredFiles = new List<string>();

            foreach (var file in rcFiles)
            {
                string fileDirectory = Path.GetDirectoryName(file);
                if (fileDirectory == null)
                    continue;

                filteredFiles.Add(file);
            }

            return filteredFiles;
        }

        public void ExportToDll()
        {

        }

        public void ExportToXML()
        {

        }
    }
}

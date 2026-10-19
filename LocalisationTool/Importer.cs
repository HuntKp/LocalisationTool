using Language;
using LocalisationTool;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

using Excel = Microsoft.Office.Interop.Excel;

namespace ImportExport
{
    internal class Importer
    {
        public MainWindow MainWindow { get; set; }

        public void ImportSingleFile()
        {
            var openFileDialog = new OpenFileDialog();
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                LoadInOneFile(openFileDialog.FileName);
            }
        }

        public void ImportAllRCFromFolder()
        {
            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = "Select a folder containing .rc files";
                dialog.ShowNewFolderButton = false;
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    string selectedPath = dialog.SelectedPath;
                    // Use the selectedPath as needed, e.g.:
                    LoadFromRC(selectedPath);
                }
            }
        }

        public void LoadInOneFile(string filePath)
        {
            string sFileName = Path.GetFileName(filePath);

            LocalisationProject languageProject = new LocalisationProject
            {
                ProjectName = sFileName,
                ProjectPath = Path.GetDirectoryName(filePath) ?? string.Empty,
                Languages = new List<LocalisationLanguage>()
            };

            ExtractDataFromRCFile(filePath, languageProject);

            MainWindow.SourceLocalisationContainer.Projects.Add(languageProject);
        }

        //load in a language file
        public void LoadFromRC(string folderPath)
        {
            List<string> rcFileList = FindRcFiles(folderPath);

            if (rcFileList.Count == 0)
            {
                Console.WriteLine("No .rc files found in the specified folder.");
                return;
            }
            else
            {
                Console.WriteLine($"Found {rcFileList.Count} .rc files in the specified folder.");

                for (int i = 0; i < rcFileList.Count; i++)
                {
                    var sFilePath = rcFileList[i];

                    string sFileName = Path.GetFileName(sFilePath);

                    LocalisationProject languageProject = new LocalisationProject
                    {
                        ProjectName = sFileName,
                        ProjectPath = Path.GetDirectoryName(sFilePath) ?? string.Empty,
                        Languages = new List<LocalisationLanguage>()
                    };

                    ExtractDataFromRCFile(sFilePath, languageProject);

                    if (languageProject.Languages.Count == 0)
                    {
                        Console.WriteLine($"No languages found in {sFileName}. Skipping this file.");
                        continue;
                    }
                    else
                    {
                        //languageProject.Languages = languageProject.Languages.OrderBy(l => l.Language).ToList();
                        var englishUK = languageProject.Languages.FirstOrDefault(l => l.Language == "English (UK)");
                        var englishUS = languageProject.Languages.FirstOrDefault(l => l.Language == "English (US)");
                        var otherLanguages = languageProject.Languages
                            .Where(l => l.Language != "English (UK)" && l.Language != "English (US)")
                            .OrderBy(l => l.Language)
                            .ToList();

                        var orderedLanguages = new List<LocalisationLanguage>();
                        if (englishUK != null)
                            orderedLanguages.Add(englishUK);
                        if (englishUS != null)
                            orderedLanguages.Add(englishUS);
                        orderedLanguages.AddRange(otherLanguages);

                        languageProject.Languages = orderedLanguages;

                        MainWindow.SourceLocalisationContainer.Projects.Add(languageProject);
                    }
                }

            }
        }

        public List<string> FindRcFiles(string folderPath)
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

        public void ExtractDataFromRCFile(string filePath, LocalisationProject languageProject)
        {
            var lines = File.ReadAllLines(filePath);
            string currentLanguage = "Unknown";
            //List<LocalisationEntry> currentEntries = null;
            var entryPattern = new Regex(@"^\s*([A-Za-z0-9_]+)\s+""(.*)""\s*$");

            if (languageProject.Languages == null)
                languageProject.Languages = new List<LocalisationLanguage>();

            LocalisationLanguage currentLocalisationLanguage = null;

            foreach (var line in lines)
            {
                var trimmed = line.Trim();

                // Detect LANGUAGE statement
                if (trimmed.StartsWith("LANGUAGE", StringComparison.OrdinalIgnoreCase))
                {
                    currentLanguage = trimmed.Substring("LANGUAGE".Length).Trim();

                    currentLocalisationLanguage = languageProject.Languages.FirstOrDefault(l => l.Language == currentLanguage);
                    if (currentLocalisationLanguage == null)
                    {
                        currentLocalisationLanguage = new LocalisationLanguage();
                        currentLocalisationLanguage.Language = MainWindow.LanguageStringMatch(currentLanguage);
                        currentLocalisationLanguage.StringTable = Path.GetFileName(filePath);
                        currentLocalisationLanguage.LocalisationEntries = new List<LocalisationEntry>();

                        languageProject.Languages.Add(currentLocalisationLanguage);
                    }

                    continue;
                }

                var match = entryPattern.Match(line);
                if (match.Success && currentLocalisationLanguage != null)
                {

                    currentLocalisationLanguage.LocalisationEntries.Add(new LocalisationEntry
                    {
                        ID = match.Groups[1].Value,
                        Text = match.Groups[2].Value,
                        Language = currentLanguage,
                        FileName = Path.GetFileName(filePath),
                        Type = "StringTable"
                    });
                }
            }
        }

        public void ImportFromExcel(string sExcelPath)
        {
            LocalisationContainer excelLanguageContainer = MainWindow.ExcelLocalisationContainer;//.Projects.Clear();
            //excelLanguageContainer.Projects.Clear();
            excelLanguageContainer.Projects = new List<LocalisationProject>();

            var excelApp = new Excel.Application();
            Excel.Workbook workbook = null;

            workbook = excelApp.Workbooks.Open(sExcelPath);

            if (workbook == null)
                return;

            foreach (Excel.Worksheet worksheet in workbook.Sheets)
            {
                //Worksheet names will represent the different LocalisationProjects
                LocalisationProject workSheetProject = new LocalisationProject();
                //Get the name from the worksheet
                workSheetProject.ProjectName = worksheet.Name;
                workSheetProject.Languages = new List<LocalisationLanguage>();

                //Read header row 1 to get the languages
                int col = 2;
                //var languageNames = new List<string>();
                while (true)
                {
                    var langCell = (worksheet.Cells[1, col] as Excel.Range).Value2;
                    if (langCell == null)
                        break;

                    LocalisationLanguage language = new LocalisationLanguage
                    {
                        Language = /*LanguageStringMatch(*/langCell.ToString()/*)*/,
                        StringTable = workSheetProject.ProjectName,
                        LocalisationEntries = new List<LocalisationEntry>()
                    };

                    //grab the languages at this point
                    int row = 2;
                    while (true)
                    {
                        var idCell = (worksheet.Cells[row, 1] as Excel.Range).Value2;
                        if (idCell == null)
                            break;
                        string id = idCell.ToString();

                        var textCell = (worksheet.Cells[row, col] as Excel.Range).Value2;
                        string text = textCell != null ? textCell.ToString() : "";

                        var entry = new LocalisationEntry
                        {
                            ID = id,
                            Text = text,
                            Language = MainWindow.LanguageStringMatch(langCell.ToString()),
                            FileName = worksheet.Name,
                            Type = "ExcelImport"
                        };

                        language.LocalisationEntries.Add(entry);

                        row++;
                    }

                    workSheetProject.Languages.Add(language);
                    col++;
                }

                // Release worksheet COM object
                Marshal.ReleaseComObject(worksheet);

                if (workSheetProject.Languages.Count() > 0)
                    excelLanguageContainer.Projects.Add(workSheetProject);

            }

            if (workbook != null)
            {
                workbook.Close(false);
                Marshal.ReleaseComObject(workbook);
            }
            excelApp.Quit();
            Marshal.ReleaseComObject(excelApp);
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;

using Language;

namespace LocalisationTool
{
    public partial class MainWindow : Window
    {
        public Language.LocalisationContainer SourceLocalisationContainer;
        public Language.LocalisationContainer ExcelLocalisationContainer;

        ImportExport.Importer importer;
        ImportExport.Exporter exporter;

        public MainWindow()
        {
            InitializeComponent();

            //Langauge Containers
            SourceLocalisationContainer = new Language.LocalisationContainer();
            SourceLocalisationContainer.ContainerName = "Main";
            SourceLocalisationContainer.Projects = new List<LocalisationProject>();

            ExcelLocalisationContainer = new Language.LocalisationContainer();
            ExcelLocalisationContainer.ContainerName = "Excel";
            ExcelLocalisationContainer.Projects = new List<LocalisationProject>();

            //import and export
            importer = new ImportExport.Importer();
            importer.MainWindow = this;

            exporter = new ImportExport.Exporter();
            exporter.MainWindow = this;
        }

        //On Import
        private void OnImportSingleStringTable(object sender, RoutedEventArgs e)
        {
            importer.ImportSingleFile();

            UpdateGridUI("Main");
        }
        private void OnImportAllRCStringTable(object sender, RoutedEventArgs e)
        {
            importer.ImportAllRCFromFolder();

            UpdateGridUI("Main");
        }

        //RC String table Import

        //DLL String table Import

        //XML Import

        //Excel import
        private async void OnImportExcel(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog();
            if (openFileDialog.ShowDialog() == true)
            {
                ShowLoading();

                await Task.Run(() =>
                {
                    importer.ImportFromExcel(openFileDialog.FileName);
                });

                HideLoading();
                // TODO: Add your file loading logic here
                MessageBox.Show($"Import file: {openFileDialog.FileName}");
            }

            UpdateGridUI("Excel");
        }

        //On Export

        //Export to RC String Tables
        private void OnExportToProject(object sender, RoutedEventArgs e)
        {
            exporter.ExportToRCFiles(ExcelLocalisationContainer);
        }

        //Export to DLL String tables

        //Export to XML

        //Export to Excel
        private async void OnExportExcel(object sender, RoutedEventArgs e)
        {
            var saveFileDialog = new Microsoft.Win32.SaveFileDialog
            {
                Filter = "Excel Files (*.xlsx)|*.xlsx",
                FileName = "Localisation.xlsx"
            };
            if (saveFileDialog.ShowDialog() == true)
            {
                ShowLoading();
                await Task.Run(() =>
                {
                    exporter.ExportToExcel_Interop(saveFileDialog.FileName, SourceLocalisationContainer.Projects);
                });
                HideLoading();

                MessageBox.Show("Exported to Excel successfully!");
            }
        }


        //TOOLS
        private void OnOptions(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Options clicked! This feature is not yet implemented.");
        }

        private void OnAbout(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("This localisation tool is still in development");
        }

        //UI Updates

        public void UpdateProjectUI(LocalisationContainer localisationContainer, ComboBox projectComboBox)
        {
            var allProjects = new LocalisationProject
            {
                ProjectName = "All Projects",
                Languages = new List<LocalisationLanguage>() // Not used directly
            };

            var projectList = new List<LocalisationProject> { allProjects };
            if (localisationContainer.Projects != null)
                projectList.AddRange(localisationContainer.Projects);

            projectComboBox.ItemsSource = projectList;
            projectComboBox.SelectedIndex = 0;
        }

        //update ui
        public void UpdateGridUI(string stab)
        {
            LocalisationContainer localisationContainer;
            DataGrid localisationDataGrid;
            ComboBox projComboBox;
            ComboBox langComboBox;

            switch (stab)
            {
                case "Excel":
                    localisationContainer = ExcelLocalisationContainer;
                    localisationDataGrid = LocalisationDataGrid_Excel;
                    projComboBox = ProjectComboBox_Excel;
                    langComboBox = LanguageComboBox_Excel;
                    break;

                case "Main":
                default:
                    localisationContainer = SourceLocalisationContainer;
                    localisationDataGrid = LocalisationDataGrid;
                    projComboBox = ProjectComboBox;
                    langComboBox = LanguageComboBox;

                    break;
            }

            projComboBox.ItemsSource = localisationContainer.Projects;
            UpdateProjectUI(localisationContainer, projComboBox);

            // Select the first language by default if available
            if (localisationContainer.Projects != null && localisationContainer.Projects.Count > 0)
            {
                if (projComboBox.SelectedItem == null)
                    projComboBox.SelectedIndex = 0;

                var selectedProject = projComboBox.SelectedItem as LocalisationProject;
                if (selectedProject != null)
                {
                    // Bind ComboBox to the list of languages
                    langComboBox.ItemsSource = selectedProject.Languages;

                    // Select the first language by default if available
                    if (selectedProject.Languages != null && selectedProject.Languages.Count > 0)
                    {
                        if (langComboBox.SelectedItem == null)
                            langComboBox.SelectedIndex = 0;

                        // Show entries for the selected language
                        var selectedLanguage = langComboBox.SelectedItem as LocalisationLanguage;
                        if (selectedLanguage != null)
                            localisationDataGrid.ItemsSource = selectedLanguage.LocalisationEntries;
                        else
                            localisationDataGrid.ItemsSource = null;
                    }
                    else
                    {
                        localisationDataGrid.ItemsSource = null;
                    }
                }
                else
                    localisationDataGrid.ItemsSource = null;
            }
            else
                localisationDataGrid.ItemsSource = null;

            switch (stab)
            {
                case "Excel":
                    ProjectComboBox_SelectionChanged_Excel(projComboBox, null);
                    break;

                case "Main":
                default:
                    ProjectComboBox_SelectionChanged(projComboBox, null);
                    break;
            }

            ProjectComboBox_SelectionChanged(projComboBox, null);
        }

        //Combo Box Selection Changed

        //Project Combo Box
        private void ProjectComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ProjectComboBox_SelectionChange(SourceLocalisationContainer, ProjectComboBox, LanguageComboBox, LocalisationDataGrid);
        }
        private void ProjectComboBox_SelectionChanged_Excel(object sender, SelectionChangedEventArgs e)
        {
            ProjectComboBox_SelectionChange(ExcelLocalisationContainer, ProjectComboBox_Excel, LanguageComboBox_Excel, LocalisationDataGrid_Excel);
        }

        private void ProjectComboBox_SelectionChange(LocalisationContainer localisationContainer, ComboBox projectComboBox, ComboBox languageComboBox, DataGrid localisationDataGrid)
        {
            var selectedProject = projectComboBox.SelectedItem as LocalisationProject;
            if (selectedProject == null)
            {
                languageComboBox.ItemsSource = null;
                localisationDataGrid.ItemsSource = null;
                return;
            }

            if (selectedProject.ProjectName == "All Projects")
            {
                // Gather all unique language names from all projects
                var allLanguages = localisationContainer.Projects
                    .SelectMany(p => p.Languages)
                    .GroupBy(l => l.Language)
                    .Select(g => g.First())
                    .ToList();

                languageComboBox.ItemsSource = allLanguages;
                languageComboBox.SelectedIndex = 0;
            }
            else
            {
                languageComboBox.ItemsSource = selectedProject.Languages;
                languageComboBox.SelectedIndex = 0;
            }

            // Manually trigger the language selection changed logic
            LanguageComboBox_SelectionChange(localisationContainer, projectComboBox, languageComboBox, localisationDataGrid);
        }

        //Language Combo Box
        private void LanguageComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            LanguageComboBox_SelectionChange(SourceLocalisationContainer, ProjectComboBox, LanguageComboBox, LocalisationDataGrid);
        }

        private void LanguageComboBox_SelectionChanged_Excel(object sender, SelectionChangedEventArgs e)
        {
            LanguageComboBox_SelectionChange(ExcelLocalisationContainer, ProjectComboBox_Excel, LanguageComboBox_Excel, LocalisationDataGrid_Excel);
        }

        private void LanguageComboBox_SelectionChange(LocalisationContainer localisationContainer, ComboBox projectComboBox, ComboBox languageComboBox, DataGrid localisationDataGrid)
        {
            var selectedProject = projectComboBox.SelectedItem as LocalisationProject;
            var selectedLanguage = languageComboBox.SelectedItem as LocalisationLanguage;

            if (selectedProject == null || selectedLanguage == null)
            {
                localisationDataGrid.ItemsSource = null;
                return;
            }

            if (selectedProject.ProjectName == "All Projects")
            {
                // Show all entries from all projects for the selected language
                var allEntries = localisationContainer.Projects
                    .SelectMany(p => p.Languages)
                    .Where(l => l.Language == selectedLanguage.Language)
                    .SelectMany(l => l.LocalisationEntries)
                    .ToList();

                localisationDataGrid.ItemsSource = allEntries;
            }
            else
            {
                localisationDataGrid.ItemsSource = selectedLanguage.LocalisationEntries;
            }
        }


        private void EntryInfo_Open(object sender, RoutedEventArgs e)
        {
            var button = sender as Button;
            var entry = button?.Tag as LocalisationEntry;
            if (entry != null)
                OpenLocalisationEntryMoreInfo(button);
        }

        private void EntryInfo_Close(object sender, RoutedEventArgs e)
        {
            LocalisationEntryMore.Visibility = Visibility.Collapsed;
        }

        public class LanguageEntryViewModel
    {
        public string Language { get; set; }
        public string Text { get; set; }
        }

        void OpenLocalisationEntryMoreInfo(Button button)
        {
            var entry = button?.Tag as LocalisationEntry;

            if (entry == null)
                return;

            // 1. Determine which tab is active
            string tabKey = null;
            if (MainTabControl.SelectedItem is TabItem selectedTab)
            {
                tabKey = selectedTab.Header as string;
            }

            LocalisationContainer container = null;
            switch (tabKey)
            {
                case "Main - Localisation Source":
                default:
                    container = SourceLocalisationContainer;
                    break;
                case "Excel - Data from Localisation Excel":
                    container = ExcelLocalisationContainer;
                    break;
            }

            // 2. Find the project that contains this entry
            LocalisationProject foundProject = null;
            LocalisationLanguage foundLanguage = null;
            foreach (var project in container.Projects)
            {
                foreach (var language in project.Languages)
                {
                    if (language.LocalisationEntries.Contains(entry))
                    {
                        foundProject = project;
                        foundLanguage = language;
                        break;
                    }
                }
                if (foundProject != null)
                    break;
            }

            LocalisationEntryMore.Visibility = Visibility.Visible;
            LocalisationEntryMore.DataContext = LocalisationDataGrid.SelectedItem as LocalisationEntry;

            if (foundProject != null)
            {
                string title = "Tab: " + container.ContainerName + "\nProject: " + foundProject.ProjectName + "\nID: " + entry.ID;

                LocalisationEntryMore_IDTitle.Text = title;

                Dictionary<string, string> languageEntriesSameID = foundProject.GetAllLanguagesForID(entry.ID);

                // Convert dictionary to a list of view models
                var gridData = languageEntriesSameID
                    .Select(kvp => new LanguageEntryViewModel { Language = LanguageStringMatch(kvp.Key), Text = kvp.Value })
                    .ToList();

                LocalisationDataGrid_Entry.ItemsSource = gridData;
            }
        }

        //Loading Screen
        public void ShowLoading()
        {
            LoadingOverlay.Visibility = Visibility.Visible;
            // Optionally force UI update
            System.Windows.Threading.Dispatcher.CurrentDispatcher.Invoke(System.Windows.Threading.DispatcherPriority.Background, new Action(delegate { }));
        }

        public void HideLoading()
        {
            LoadingOverlay.Visibility = Visibility.Collapsed;
        }

        //Language Names from String Tables
        Dictionary<string, string> LanguageNames = new Dictionary<string, string>
        {
            { "LANG_ENGLISH, SUBLANG_ENGLISH_UK", "English (UK)" },
            { "LANG_ENGLISH, SUBLANG_ENGLISH_US", "English (US)" },
            { "LANG_WELSH, SUBLANG_DEFAULT", "Welsh" },
            { "LANG_IRISH, SUBLANG_IRISH_IRELAND", "Irish" },
            { "LANG_SCOTTISH_GAELIC, SUBLANG_DEFAULT", "Scottish Gaelic" },
            { "LANG_NORWEGIAN, SUBLANG_NEUTRAL", "Norwegian" },
            { "LANG_SWEDISH, SUBLANG_SWEDISH", "Swedish" },
            { "LANG_DANISH, SUBLANG_DEFAULT", "Danish" },
            { "LANG_FINNISH, SUBLANG_DEFAULT", "Finnish" },
            { "LANG_DUTCH, SUBLANG_DUTCH_BELGIAN", "Dutch" },
            { "LANG_FRENCH, SUBLANG_FRENCH", "French" },
            { "LANG_GERMAN, SUBLANG_GERMAN", "German" },
            { "LANG_SPANISH, SUBLANG_SPANISH", "Spanish" },
            { "LANG_SPANISH, SUBLANG_SPANISH_MODERN", "Spanish" },
            { "LANG_PORTUGUESE, SUBLANG_PORTUGUESE", "Portuguese" },
            { "LANG_ITALIAN, SUBLANG_ITALIAN", "Italian" },
            { "LANG_GREEK, SUBLANG_DEFAULT", "Greek" },
            { "LANG_CZECH, SUBLANG_DEFAULT", "Czech" },
            { "LANG_POLISH, SUBLANG_DEFAULT", "Polish" },
            { "LANG_LITHUANIAN, SUBLANG_LITHUANIAN", "Lithuanian" },
            { "LANG_ALBANIAN, SUBLANG_DEFAULT", "Albanian" },
            { "LANG_BULGARIAN, SUBLANG_DEFAULT", "Bulgarian" },
            { "LANG_TURKISH, SUBLANG_DEFAULT", "Turkish" },
            { "LANG_UKRAINIAN, SUBLANG_DEFAULT", "Ukrainian" },
            { "LANG_RUSSIAN, SUBLANG_DEFAULT", "Russian" },
            { "LANG_JAPANESE, SUBLANG_NEUTRAL", "Japanese" },
            { "LANG_JAPANESE, SUBLANG_DEFAULT", "Japanese" },
            { "LANG_KOREAN, SUBLANG_NEUTRAL", "Korean" },
            { "LANG_CHINESE, SUBLANG_CHINESE_SIMPLIFIED", "Chinese (Simplified)" },
            { "LANG_CHINESE, SUBLANG_CHINESE_TRADITIONAL", "Chinese (Traditional)" },
            { "LANG_NEUTRAL, SUBLANG_NEUTRAL", "Neutral - NOT A LANGUAGE" }
        };

        public string LanguageStringMatch(string sLanguage)
        {
            string sResult = sLanguage;

            if (LanguageNames.ContainsKey(sLanguage))
            {
                LanguageNames.TryGetValue(sLanguage, out sResult);
            }
            else
            {
                //if sLanguage isn't the key, it may be the value, return the key.
                foreach (var entry in LanguageNames)
                {
                    string sValue;
                    LanguageNames.TryGetValue(entry.Key, out sValue);

                    if (sValue.Equals(sLanguage))
                    {
                        sResult = entry.Key;
                        break;
                    }
                }
            }

            return sResult;
        }
    }
}

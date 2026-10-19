using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Language
{
    internal class LanguageController
    {
        //add a brand new one, this will either be a copy past of any languages we already have, or slightly smarter and try to translate entries automatically
        public void AddNewLanguage(LocalisationContainer localisationContainer)
        {
            string newLanguageName = "New Language"; // Placeholder for the new language name
            string newStringTable = "New String Table"; // Placeholder for the new string table name

            if (localisationContainer.Projects == null || localisationContainer.Projects.Count == 0)
                return;

            foreach (var localisationProject in localisationContainer.Projects)
            {
                if (localisationProject.Languages == null || localisationProject.Languages.Count == 0)
                    return;

                // Copy entries from the first language
                var firstLanguage = localisationProject.Languages[0];

                // Deep copy the entries
                var copiedEntries = firstLanguage.LocalisationEntries
                    .Select(entry => new LocalisationEntry
                    {
                        ID = entry.ID,
                        Text = entry.Text,
                        Language = entry.Language,
                        FileName = entry.FileName,
                        Type = entry.Type
                    })
                    .ToList();

                // Create the new language (you may want to set Language and StringTable as needed)
                var newLanguage = new LocalisationLanguage
                {
                    Language = newLanguageName,
                    StringTable = newStringTable,
                    LocalisationEntries = copiedEntries
                };

                localisationProject.Languages.Add(newLanguage);
            }

        }

        //Removes a language, potentially a dangerous option
        public void DeleteLanguage(LocalisationContainer localisationContainer)
        {
            string languageToDelete = "Language to Delete"; // Placeholder for the language name to delete

            if (localisationContainer.Projects == null || localisationContainer.Projects.Count == 0)
                return;

            foreach (var localisationProject in localisationContainer.Projects)
            {
                for (int i = localisationProject.Languages.Count - 1; i >= 0; i--)
                {
                    if (localisationProject.Languages[i].Language == languageToDelete)
                    {
                        localisationProject.Languages.RemoveAt(i);
                        break; // Exit the loop after removing the language
                    }
                }
            }
        }

        //Add a new localisation entry
        public void AddNewLocalisationEntry(LocalisationContainer localisationContainer)
        {
            //huntkp - set this to actually be the currentl selected project
            LocalisationProject currentSelectedProject = localisationContainer.Projects.FirstOrDefault(); // Assuming you have a way to get the currently selected project

            if (currentSelectedProject == null)
                return;

            if (currentSelectedProject.Languages == null || currentSelectedProject.Languages.Count == 0)
                return;

            var newLocalisationEntry = new LocalisationEntry
            {
                ID = "NewId", // Placeholder for the new entry ID
                Text = "New Text", // Placeholder for the new entry text
                Language = currentSelectedProject.Languages[0].Language, // Use the first language as default
                FileName = "NewFileName", // Placeholder for the file name
                Type = "NewType" // Placeholder for the type
            };

            for (int i = 0; i < currentSelectedProject.Languages.Count(); i++)
            {
                // Add the new entry to each language's localisation entries
                currentSelectedProject.Languages[i].LocalisationEntries.Add(newLocalisationEntry);
            }
        }

        //Remove a localisation entry
        public void RemoveLocalisationEntry(LocalisationContainer localisationContainer)
        {
            //huntkp - set this to actually be the currentl selected project
            LocalisationProject currentSelectedProject = localisationContainer.Projects.FirstOrDefault(); // Assuming you have a way to get the currently selected project

            if (currentSelectedProject == null)
                return;

            if (currentSelectedProject.Languages == null || currentSelectedProject.Languages.Count == 0)
                return;

            string idtoremove = "10";// Placeholder for the ID to remove

            for (int i = 0; i < currentSelectedProject.Languages.Count(); i++)
            {
                for (int j = 0; j < currentSelectedProject.Languages[i].LocalisationEntries.Count; j++)
                {
                    if (currentSelectedProject.Languages[i].LocalisationEntries[j].ID == idtoremove)
                    {
                        currentSelectedProject.Languages[i].LocalisationEntries.RemoveAt(j);
                        break; // Exit the loop after removing the entry
                    }
                }
            }
        }

        //Edit a localisation entry, most likely only the body of the text will be altered, but we may allow the ids to be edited, but these would have to be synced across all languages
        public void ChangeLocalisationEntry()
        {

        }
    }
}

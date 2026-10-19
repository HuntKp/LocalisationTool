using System.Collections.Generic;

namespace Language
{
    public class LocalisationProject
    {
        public string ProjectName { get; set; }
        public string ProjectPath { get; set; }
        public List<LocalisationLanguage> Languages { get; set; }

        public Dictionary<string, string> GetAllLanguagesForID(string id)
        {
            Dictionary<string, string> idLanguageEntities = new Dictionary<string, string>();

            foreach (var language in Languages)
            {
                foreach (var entry in language.LocalisationEntries)
                {
                    if (string.Equals(id, entry.ID))
                    {
                        idLanguageEntities.Add(entry.Language, entry.Text);
                    }
                }
            }

            return idLanguageEntities;
        }
    }
}

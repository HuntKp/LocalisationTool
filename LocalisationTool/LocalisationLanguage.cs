using System.Collections.Generic;

namespace Language
{
    public class LocalisationLanguage
    {
        //the language of this localisation list
        public string Language { get; set; }
        //the file we got the information from, this may be removed
        public string StringTable { get; set; }
        //A list of all of our localisation entries, this presents our string table
        public List<LocalisationEntry> LocalisationEntries { get; set; }
    }
}

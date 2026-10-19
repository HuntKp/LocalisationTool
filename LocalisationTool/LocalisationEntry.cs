namespace Language
{
    public class LocalisationEntry
    {
        //the value id, this will link between languages
        public string ID { get; set; }
        ////the value id, this will link between languages
        //public string Value { get; set; }
        //The body of the text
        public string Text { get; set; }
        //The language this entry is in
        public string Language { get; set; }
        //The name of the file which we got the information from
        public string FileName { get; set; }
        //the type will lets us know where the information is from and where it will go, is it stringtable, string literals, etc 
        public string Type { get; set; }
    }
}

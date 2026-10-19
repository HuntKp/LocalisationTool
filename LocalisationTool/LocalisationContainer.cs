using System.Collections.Generic;

namespace Language
{
    public class LocalisationContainer
    {
        public string ContainerName { get; set; }
        //store each project, and each project will have the languages
        public List<LocalisationProject> Projects { get; set; }
    }
}

using System.Collections.Generic;

namespace SharePointAPI.Models
{
    public class DocumentModel
    {
        public string _id{ get; set; }
        public string file_url { get; set; }
        public string filename { get; set; }
        public string foldername { get; set; }
        public string sitecontent { get; set; }
        public string list { get; set; }
        public string site { get; set; }
        public IDictionary<string, string> fields { get; set; }
        public IDictionary<string, string> taxFields { get; set; }
        public IDictionary<string, List<string>> taxListFields { get; set; }

    }
}
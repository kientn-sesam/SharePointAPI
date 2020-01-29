namespace SharePointAPI.Models
{
    public class SMBCredential
    {
        public string username { get; set; }

        public string password { get; set; }
        
        public string domain { get; set; }
        public string ipaddr { get; set; }
        public string share { get; set; }
    }
}
using System.Collections.Generic;
using System;

namespace SharePointAPI.Models
{
    public class TermModel
    {
        
        public Guid Id { get; set; }
        public string Name { get; set; }
        public string Description { get; set; }
        public List<Lbl> Labels { get; set; }
        
    }
    public class Lbl 
    {
        public bool IsDefaultForLanguage { get; set; }
        public int Language { get; set; }
        public string Value { get; set; }
    }
}
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using Microsoft.SharePoint.Client.UserProfiles;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using SharePointAPI.Models;

namespace SharePointAPI.Middleware
{
    public class TermHelper
    {
        /// <summary>
        /// Helper Methods to get a TermId by a Name
        /// </summary>
        /// <param name="cc">The Authenticated ClientContext</param>
        /// <param name="term">The Term Name do lookup.</param>
        /// <param name="termSetId">The TermSet Guid</param>
        /// <returns></returns>
        public static string GetTermIdByName(ClientContext cc, string term, Guid termSetId)
        {
            string _resultTerm = string.Empty;

            var _taxSession = TaxonomySession.GetTaxonomySession(cc);
            var _termStore = _taxSession.GetDefaultSiteCollectionTermStore();
            var _termSet = _termStore.GetTermSet(termSetId);

            var _termMatch = new LabelMatchInformation(cc)
            {
               Lcid = 1033,
               TermLabel = term,     
               TrimUnavailable = true
            };

            var _termCollection = _termSet.GetTerms(_termMatch);
            cc.Load(_taxSession);
            cc.Load(_termStore);
            cc.Load(_termSet);
            cc.Load(_termCollection);
            cc.ExecuteQuery();

            if (_termCollection.Count() > 0)
                _resultTerm = _termCollection.First().Id.ToString();

            return _resultTerm;

        }

        public static TermModel Term(Term term)
        {
            List<Lbl> labels = new List<Lbl>();
            LabelCollection labelCollection = term.Labels;
            foreach (var label in labelCollection)
            {
                labels.Add(new Lbl(){IsDefaultForLanguage = label.IsDefaultForLanguage, Language = label.Language, Value = label.Value});
            }

            return new TermModel{ Id = term.Id, Name = term.Name, Description = term.Description,Labels = labels };

        }

    }
}
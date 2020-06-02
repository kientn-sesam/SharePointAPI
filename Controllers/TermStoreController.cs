using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using SharePointAPI.Models;
using SharePointAPI.Middleware;

namespace SharePointAPI.Controllers
{
    [ApiController]
    [Produces("application/json")]
    [Route("api/[controller]/[action]")]
    public class TermStoreController : ControllerBase
    {
        private readonly ILogger<TermStoreController> _logger;
        public string _username, _password, _baseurl;
        public TermStoreController(ILogger<TermStoreController> logger)
        {
            _logger = logger;
            
            if(Environment.GetEnvironmentVariable("baseurl") != null){
                Console.WriteLine("Baseline url: " + Environment.GetEnvironmentVariable("baseurl"));
                _baseurl = Environment.GetEnvironmentVariable("baseurl");
                _username = Environment.GetEnvironmentVariable("username");
                _password = Environment.GetEnvironmentVariable("password");

            }
            else{
                using (var file = System.IO.File.OpenText("test.json"))
                {
                    var reader = new JsonTextReader(file);
                    var jObject = JObject.Load(reader);
                    _baseurl = jObject.GetValue("baseurl").ToString();
                    _username = jObject.GetValue("username").ToString();
                    _password = jObject.GetValue("password").ToString();
                }
            }
            
        }

        [HttpGet]
        [Produces("application/json")]
        [Consumes("application/json")]
        /// <summary>
        /// Fetch all terms from term set
        /// Required query termsstore guid, termgroup and term set.
        /// 
        /// GET
        /// /api/termstore/terms?termstore=244732abfd89492bab7a195786b2cdc7&termgroup=AE-SPOR
        /// </summary>
        /// <param name=""termstore""></param>
        /// <returns></returns>

        public async Task<IActionResult> Terms([FromQuery(Name = "termstore")] string _termstore, [FromQuery(Name = "termgroup")] string _termgroup)
        {

            using(ClientContext cc = AuthHelper.GetClientContextForUsernameAndPassword(_baseurl, _username, _password))
            try
            {
                cc.RequestTimeout = -1;

                List<TermModel> listTerms = new List<TermModel>();
                TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(cc);
                var termStores = taxonomySession.TermStores;
                TermStore termStore = termStores.GetById(new Guid(_termstore));
                var termGroups = termStore.Groups;
                var termGroup = termGroups.GetByName(_termgroup);
                var termSets = termGroup.TermSets;
                cc.Load(termGroup,
                               group => group.Id,
                               group => group.Name,
                               group => group.TermSets.Include(
                                   set => set.Name,
                                   set => set.Id,
                                   set => set.Terms.Include(
                                       term => term.Id,
                                       term => term.Name,
                                       term => term.Description,
                                       term => term.Labels,
                                term  => term.Terms.Include(
                                    term => term.Name,
                                    term => term.Description,
                                    term => term.Id,
                                    term=> term.Labels,
                                    term => term.Terms.Include(
                                        term => term.Name,
                                        term => term.Description,
                                        term=> term.Id,
                                        term => term.Labels,
                                        term => term.Terms.Include(
                                            term => term.Name,
                                            term => term.Description,
                                            term=> term.Id,
                                            term=> term.Labels,
                                            term => term.Terms.Include(
                                                term => term.Name,
                                                term => term.Description,
                                                term=> term.Id,
                                                term=> term.Labels,
                                                term => term.Terms.Include(
                                                    term => term.Name,
                                                    term => term.Description,
                                                    term=> term.Id,
                                                    term=> term.Labels
                                                )
                                            
                                        )
                                    )
                                )
                            )
                           )
                       )
                );


                await cc.ExecuteQueryAsync();
                for (int i = 0; i < termSets.Count; i++)
                {

                        TermSet termSet = termSets[i];

                        var terms = termSet.Terms;
                        foreach (var Aterm in terms)
                        {   
                            listTerms.Add(TermHelper.Term(Aterm));
                            TermCollection BTerms = Aterm.Terms;
                            if (BTerms.Count > 0)
                            {
                                foreach (var BTerm in BTerms)
                                {
                                    listTerms.Add(TermHelper.Term(BTerm));
                                    TermCollection CTerms = BTerm.Terms;
                                    if (CTerms.Count > 0)
                                    {
                                        foreach (var CTerm in CTerms)
                                        {
                                            listTerms.Add(TermHelper.Term(CTerm));
                                            TermCollection DTerms = CTerm.Terms;
                                            if (DTerms.Count > 0)
                                            {
                                                foreach (var DTerm in DTerms)
                                                {
                                                    listTerms.Add(TermHelper.Term(DTerm));
                                                    TermCollection ETerms = DTerm.Terms;
                                                    if (ETerms.Count > 0)
                                                    {
                                                        foreach (var ETerm in ETerms)
                                                        {
                                                            listTerms.Add(TermHelper.Term(ETerm));
                                                            TermCollection FTerms = ETerm.Terms;
                                                            if (FTerms.Count > 0)
                                                            {
                                                                foreach (var FTerm in FTerms)
                                                                {
                                                                    listTerms.Add(TermHelper.Term(FTerm));
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }

                                }
                            }
                        }
                        

                }
                    
                
                
                return new OkObjectResult(listTerms);

                
            }
            catch (System.Exception)
            {
                
                throw;
            }

        }


    }
}
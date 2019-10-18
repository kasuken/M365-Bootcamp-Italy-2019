using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using TeamGovernance.API.Models;

// For more information on enabling Web API for empty projects, visit https://go.microsoft.com/fwlink/?LinkID=397860

namespace TeamGovernance.API.Controllers
{
    [Route("api/[controller]")]
    public class TeamGovernanceController : ControllerBase
    {
        // GET: api/<controller>
        [HttpGet]
        public async Task<TeamCreationResponse> Get(string siteTitle, string sitePrefix, string owner)
        {
            var teamCreationResponse = new TeamCreationResponse();

            var graphService = new GraphService();

            var domain = owner.Split('@')[1];
            var siteAlias = CreateSiteAlias(sitePrefix, siteTitle);

            var groupId = await graphService.CreateGroup(domain, siteTitle, siteAlias, owner);
            var team = await graphService.CreateTeamFromGroup(domain, groupId);

            return teamCreationResponse;
        }

        private string CreateSiteAlias(string prefix, string siteTitle)
        {
            var teamName = prefix + siteTitle.Replace(" ", "");

            return teamName;
        }
    }
}

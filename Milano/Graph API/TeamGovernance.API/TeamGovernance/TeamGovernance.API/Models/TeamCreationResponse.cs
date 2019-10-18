using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace TeamGovernance.API.Models
{
    public class TeamCreationResponse
    {

        public string OfficeGroupName { get; set; }

        public string TeamName { get; set; }

        public string SharePointSiteUrl { get; set; }

        public string TeamUrl { get; set; }

        public string ErrorMessage { get; set; }

    }
}

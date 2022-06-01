using System;
using System.Collections.Generic;
using System.Text;

namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.AdaptiveCard
{
    public class Mentioned
    {
        public string id { get; set; }

        public string name { get; set; }
    }

    public class Entity
    {
        public string type { get; set; }

        public string text { get; set; }

        public Mentioned mentioned { get; set; }
    }
}

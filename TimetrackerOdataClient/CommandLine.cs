using System;
using System.Collections.Generic;
using System.Dynamic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CommandLine;
using CommandLine.Text;

namespace TimetrackerOdataClient
{
    public class CommandLineOptions
    {
        [Value( 0, Required = true, HelpText = "Service URL for Timetracker OData endpoint" )]
        public Uri ServiceUri { get; set; }


        [Option( 'w', Default = false, HelpText = "On-premise usage (NTLM auth)" )]
        public bool IsWindowsAuth { get; set; }

        [Option( 't', HelpText = "Token for Timetracker API (VSTS usage)" )]
        public string Token { get; set; }

        [Option( 'f', HelpText = "TFS URL with collection part" )]
        public string TfsUrl { get; set; }

        [Option( 'v', HelpText = "VSTS personal token" )]
        public string VstsToken { get; set; }

        [Option("from", HelpText = "Start date, format \"yyyy-mm-dd\".")]
        public string StartDate { get; set; }

        [Option("to", HelpText = "End date, format \"yyyy-mm-dd\".")]
        public string EndDate { get; set; }

        [Option("open", Default = false, HelpText = "Open file after generation")]
        public bool OpenFileAfterGeneration { get; set; }
    }
}
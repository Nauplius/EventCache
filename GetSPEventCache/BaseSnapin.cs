using System.Collections.Generic;
using System.ComponentModel;
using System.Management.Automation;

namespace GetSPEventCache
{
    [RunInstaller(true)]
    // ReSharper disable once InconsistentNaming
    public class SPEventCache : PSSnapIn
    {
        /// <summary>
        /// Creates an instance of CustomPSSnapInTest class. 
        /// </summary>
        public SPEventCache()
            : base()
        {
        }

        public override string Name
        {
            get { return "SPEventCache"; }
        }

        public override string Vendor
        {
            get { return "Nauplius"; }
        }

        public override string VendorResource
        {
            get
            {
                return "SPEventCache,Nauplius (Trevor Seward)";
            }
        }

        /// <summary>
        /// Specify a description of the custom PowerShell snap-in. 
        /// </summary>
        public override string Description
        {
            get
            {
                return "Provides information on SharePoint EventCache.";
            }
        }
    }
}

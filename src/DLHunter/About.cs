using System;
using System.Collections.Generic;
using System.Deployment.Application;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DLHunter
{
    class About
    {

        public static string Version()
        {
            return ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString();
        }
    }
}

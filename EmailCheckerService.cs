using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;

namespace EmailCheckerService
{
    public partial class EmailCheckerService : ServiceBase
    {
        public EmailCheckerService()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            Monitor m = new Monitor();
            m.Run();

        }

        internal void OnDebug()
        {
            OnStart(null);
        }

        protected override void OnStop()
        {
        }
    }
}

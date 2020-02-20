using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EmailTester
{
    public class Result
    {
        public string emailAddress { get; set; }
        public string message { get; set; }

        public Result()
        {

        }

        public Result(string emailAddress, string message)
        {
            this.emailAddress = emailAddress;
            this.message = message;
        }
    }
}

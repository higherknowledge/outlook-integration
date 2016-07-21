using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.Serialization.Json;
using System.Runtime.Serialization;

namespace HigherKnowledge_addin
{
    [DataContract]
    class Template
    {
        [DataMember(Name = "Subject")]
        public string subject { get; set; }

        [DataMember(Name = "CC")]
        public string cc { get; set; }

        [DataMember(Name = "Body")]
        public string[] body { get; set; }
    }
}

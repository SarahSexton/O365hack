using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web;

namespace DXNextHackatonWeb.Models
{
    public class EventModel
    {
        public Guid objectId { get; set; }
        public Organizer Organizer { get; set; }
        public DateTime Start { get; set; }
        public DateTime End { get; set; }

        public string Subject { get; set; }
    }
}
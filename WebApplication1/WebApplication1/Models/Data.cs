using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace WebApplication1.Models
{
    public class Data
    {
       
        public int userId { get; set; }
        public int id { get; set; }
        [Required]
        public string title { get; set; }
        public string body { get; set; }
        public int postId { get; set; }
        public string  email { get; set; }
        public string  name { get; set; }

    }
}
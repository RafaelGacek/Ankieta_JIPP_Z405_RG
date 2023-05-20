using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace AnkietaProjekt.Models
{
    [Table("Questions")]

    public class Question
    {
        [Key]
        public int ID { get; set; }
        public string Pytanie { get; set; }

    }
}
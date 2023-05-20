using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.ComponentModel.DataAnnotations.Schema;
using System.ComponentModel.DataAnnotations;

namespace AnkietaProjekt.Models
{
    [Table("Tests")]
    public class Test
    {
        [Key]
        public int ID { get; set; }
        [Display(Name = "Nazwa Użytkownika")]
        public String Username { get; set; }
        [Display(Name = "Płeć")]
        public Plec plec { get; set; }
        public String pytanie1 { get; set; }
        public Odpowiedz odp1 { get; set; }
        public String pytanie2 { get; set; }
        public Odpowiedz odp2 { get; set; }
        public String pytanie3 { get; set; }
        public Odpowiedz odp3 { get; set; }
        public String pytanie4 { get; set; }
        public Odpowiedz odp4 { get; set; }
        public String pytanie5 { get; set; }
        public Odpowiedz odp5{ get; set; }
        public String pytanie6 { get; set; }
        public Odpowiedz odp6 { get; set; }
        public String pytanie7 { get; set; }
        public Odpowiedz odp7 { get; set; }
        public String pytanie8 { get; set; }
        public Odpowiedz odp8 { get; set; }
        public String pytanie9 { get; set; }
        public Odpowiedz odp9 { get; set; }
        public String pytanie10 { get; set; }
        public Odpowiedz odp10 { get; set; }
        public String pytanie11 { get; set; }
        public Odpowiedz odp11 { get; set; }
        public String pytanie12 { get; set; }
        public Odpowiedz odp12 { get; set; }
        public String pytanie13 { get; set; }
        public Odpowiedz odp13 { get; set; }
        public String pytanie14 { get; set; }
        public Odpowiedz odp14 { get; set; }
        public String pytanie15 { get; set; }
        public Odpowiedz odp15 { get; set; }
        public String pytanie16 { get; set; }
        public Odpowiedz odp16 { get; set; }
        public String pytanie17 { get; set; }
        public Odpowiedz odp17 { get; set; }
        public String pytanie18 { get; set; }
        public Odpowiedz odp18 { get; set; }
        public String pytanie19 { get; set; }
        [Display(Name = "odp19")]
        public Odpowiedz odp119 { get; set; }
        public String pytanie20 { get; set; }
        public Odpowiedz odp20 { get; set; }

    }
    public enum Plec
    {
        kobieta,
        mezczyzna
    }
    public enum Odpowiedz
    {
        tak,
        nie,
        niewiem
    }
}
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using AnkietaProjekt.Models;
using System.Data.Entity.ModelConfiguration.Conventions;
using System.Data.Entity;

namespace AnkietaProjekt.DAL
{
    public class QuestionContext : DbContext
    {

        public QuestionContext() : base("DefaultConnection")
        {

        }
        public DbSet<Question> Questions { get; set; }
        public DbSet<Profile> Profiles { get; set; }
        public DbSet<Test> Testy { get; set; }




        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            modelBuilder.Conventions.Remove<PluralizingTableNameConvention>();

        }
    }
}
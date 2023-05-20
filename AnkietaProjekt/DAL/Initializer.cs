using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data.Entity;
using Microsoft.AspNet.Identity.EntityFramework;
using Microsoft.AspNet.Identity;
using AnkietaProjekt.Models;

namespace AnkietaProjekt.DAL
{
    public class Initializer : DropCreateDatabaseIfModelChanges<QuestionContext>
    {

        protected override void Seed(QuestionContext context)
        {
            var roleManager = new RoleManager<IdentityRole>(new RoleStore<IdentityRole>(new ApplicationDbContext()));
            var userManager = new UserManager<ApplicationUser>(new UserStore<ApplicationUser>(new ApplicationDbContext()));

            roleManager.Create(new IdentityRole("Admin"));

            var user = new ApplicationUser { UserName ="xena2021@int.pl" };
            string password = "Admin2023!";

            userManager.Create(user, password);

            userManager.AddToRole(user.Id, "Admin");

        }
    }
}
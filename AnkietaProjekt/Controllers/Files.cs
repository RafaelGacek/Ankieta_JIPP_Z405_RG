using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;

namespace AnkietaProjekt.Controllers
{
    public class Files
    {
        String Path { get; set; }

        public Files(string path)
        {
            Path = path;
        }
        public void RemoveFile()
        {

            if (File.Exists($@"{Path}"))
            {
                File.Delete($@"{Path}");
                //ViewBag.deleteSuccess = "true";
            }
        }


    }
}
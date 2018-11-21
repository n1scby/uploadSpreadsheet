using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using UploadSpreadsheet.Models;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using System.IO;
using OfficeOpenXml;

namespace UploadSpreadsheet.Controllers
{

    public class HomeController : Controller
    {
        private readonly IHostingEnvironment _hostingEnvironment;

        public HomeController (IHostingEnvironment hostingEnvironment)
        {
            _hostingEnvironment = hostingEnvironment;
        }


        public IActionResult Index()
        {
            List<DogData> dogs = new List<DogData>();
            return View(dogs);
        }

        public IActionResult About()
        {
            ViewData["Message"] = "Your application description page.";

            return View();
        }

        public IActionResult Contact()
        {
            ViewData["Message"] = "Your contact page.";

            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        public async Task<IActionResult> Upload(IFormFile file)
        {
            if (file == null)
            {
                return RedirectToAction(nameof(Index));
            }
            List<DogData> dogs = new List<DogData>();
            
            using(var memoryStream = new MemoryStream())
            {
                await file.CopyToAsync(memoryStream).ConfigureAwait(false);
                using(var package = new ExcelPackage(memoryStream))
                {
                    var worksheet = package.Workbook.Worksheets[0];
                    if (worksheet.Dimension.Rows > 0 && worksheet.Dimension.Columns > 0)
                    {
                        for (int row = 1; row <= worksheet.Dimension.Rows; row++)
                        {
                            DogData dog = new DogData();
                            dog.Name = worksheet.Cells[row, 1].Value.ToString();
                            dog.Breed = worksheet.Cells[row, 2].Value.ToString();
                            dog.Age = int.Parse(worksheet.Cells[row, 3].Value.ToString());
                            dog.Gender = worksheet.Cells[row, 4].Value.ToString();
                            dogs.Add(dog);

                        }
                    }
                }
            }
            return RedirectToAction(nameof(Index), dogs);
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}

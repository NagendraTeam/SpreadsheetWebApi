using Microsoft.AspNetCore.Hosting.Server;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Net.Http.Headers;
using SampleSpreadSheetApp.model;
using System;
using System.Collections.Generic;
using System.IO;

namespace SampleSpreadSheetApp.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class WorkbookController : ControllerBase
    {
        SpreadSheetDataLayer dataLayer = new SpreadSheetDataLayer();
        public WorkbookController() { }

        [Produces("application/json")]
        [HttpPost("InsertChild1Data")]
        public IActionResult InsertChild1Data(List<SpreadSheetModel> model)
        {
            dataLayer.InsertChild1Data(model, "Child1");
            return Ok();
        }
        [Produces("application/json")]
        [HttpPost("InsertChild2Data")]
        public IActionResult InsertChild2Data(List<SpreadSheetModel> model)
        {
            dataLayer.InsertChild2Data(model, "Child2");
            return Ok();
        }
        [Produces("application/json")]
        [HttpGet("GetParentWorkbookData")]
        public IActionResult GetParentWorkbookData()
        {
            var data = dataLayer.GetParentSheetData();
            return Ok(data);
        }
        [Produces("application/json")]
        [HttpGet("GetChild1WorkbookData")]
        public IActionResult GetChild1WorkbookData()
        {
            var data = dataLayer.GetChild1SheetData();
            return Ok(data);
        }
       
        [HttpGet("GetChild2WorkbookData"), DisableRequestSizeLimit]
        public IActionResult GetChild2WorkbookData()
        {
            var data = dataLayer.GetChild2SheetData();
            return Ok(data);
        }

        [Produces("application/json")]
        [HttpPost("UploadChild1Data")]
        public IActionResult UploadChild1Data()
        {
            try
            {
                var file = Request.Form.Files[0];

                var folderName = Path.Combine("assets", "files");
                var pathToSave = Path.Combine(Directory.GetCurrentDirectory(), folderName);

                if (file.Length > 0)
                {
                    var fileName = ContentDispositionHeaderValue.Parse(file.ContentDisposition).FileName.ToString();
                    var fullPath = Path.Combine(pathToSave, fileName);
                    var dbPath = Path.Combine(folderName, fileName);
                    using (var stream = new FileStream(fullPath, FileMode.Create))
                    {
                        file.CopyTo(stream);
                    }
                }
                else
                {
                    return BadRequest();
                }
            }
            catch(Exception ex)
            {

            }
            return Ok();
        }
        [HttpGet("getFile"), DisableRequestSizeLimit]
        public IActionResult getFile()
        {
            var folderName = Path.Combine("assets", "files");
            var pathToSave = Path.Combine(Directory.GetCurrentDirectory(), folderName);
            var fullPath = Path.Combine(pathToSave, "ParentWorkbookData.xlsx");

            HttpContext.Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            FileContentResult result = new FileContentResult(System.IO.File.ReadAllBytes(fullPath),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            {
                FileDownloadName = "otherfile"
            };
            return result;

            //return Ok(reader);
        }
        [HttpGet("getChild1File"), DisableRequestSizeLimit]
        public IActionResult getChild1File()
        {
            var folderName = Path.Combine("assets", "files");
            var pathToSave = Path.Combine(Directory.GetCurrentDirectory(), folderName);
            var fullPath = Path.Combine(pathToSave, "Child1WorkbookData.xlsx");

            HttpContext.Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            FileContentResult result = new FileContentResult(System.IO.File.ReadAllBytes(fullPath),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            {
                FileDownloadName = "otherfile"
            };
            return result;

            //return Ok(reader);
        }
        [HttpGet("getChild2File"), DisableRequestSizeLimit]
        public IActionResult getChild2File()
        {
            var folderName = Path.Combine("assets", "files");
            var pathToSave = Path.Combine(Directory.GetCurrentDirectory(), folderName);
            var fullPath = Path.Combine(pathToSave, "Child2WorkbookData.xlsx");

            HttpContext.Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            FileContentResult result = new FileContentResult(System.IO.File.ReadAllBytes(fullPath),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            {
                FileDownloadName = "otherfile"
            };
            return result;

            //return Ok(reader);
        }

    }
}

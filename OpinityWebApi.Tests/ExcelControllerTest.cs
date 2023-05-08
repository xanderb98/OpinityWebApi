using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Http.HttpResults;
using Microsoft.AspNetCore.Mvc;
using Moq;
using OpinityWebApi.Controllers;
using OpinityWebApi.Services;
using System;
using System.IO;
using System.Text;

namespace OpinityWebApi.Tests
{
    public class ExcelControllerTest
    {
        [Fact]
        public async Task UploadExcel_ValidFile_ReturnsOk()
        {
            var mockConvertExcelToObject = new Mock<IConvertExcelToObject>();
            var controller = new ExcelController(mockConvertExcelToObject.Object);
            
            string fileName = Path.GetFullPath(Directory.GetParent(Directory.GetCurrentDirectory()).Parent.Parent + @"\test.xlsx");
            var content = File.OpenRead(fileName);
            var file = new Mock<IFormFile>();

            file.Setup(f => f.OpenReadStream()).Returns(content);
            file.Setup(f => f.FileName).Returns("test.xlsx");
            file.Setup(f => f.Length).Returns(content.Length);

            var result = await controller.UploadExcel(file.Object);

            Assert.IsType<OkObjectResult>(result);
        }

        [Fact]
        public async Task UploadExcel_NoFile_ReturnsBadRequest()
        {
            var mockConvertExcelToObject = new Mock<IConvertExcelToObject>();
            var controller = new ExcelController(mockConvertExcelToObject.Object);
            
            var result = await controller.UploadExcel(null);

            var badRequestResult = Assert.IsType<BadRequestObjectResult>(result);

            Assert.Equal("There was no file selected", badRequestResult.Value);
        }

        [Fact]
        public async Task UploadExcel_NoValidFile_ReturnsBadRequest()
        {
            var mockConvertExcelToObject = new Mock<IConvertExcelToObject>();
            var controller = new ExcelController(mockConvertExcelToObject.Object);
            string fileName = Path.GetFullPath(Directory.GetParent(Directory.GetCurrentDirectory()).Parent.Parent + @"\test.txt");
            var content = File.OpenRead(fileName);
            var file = new Mock<IFormFile>();

            file.Setup(f => f.OpenReadStream()).Returns(content);
            file.Setup(f => f.FileName).Returns("test.txt");
            file.Setup(f => f.Length).Returns(content.Length);

            var result = await controller.UploadExcel(file.Object);

            var badRequestResult = Assert.IsType<BadRequestObjectResult>(result);

            Assert.Equal("Format does not match .xlsx or .xls extensions", badRequestResult.Value);
        }
    }
}
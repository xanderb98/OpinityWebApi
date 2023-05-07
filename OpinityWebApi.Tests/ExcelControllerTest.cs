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
            // Arrange
            var mockConvertExcelToObject = new Mock<IConvertExcelToObject>();
            var controller = new ExcelController(mockConvertExcelToObject.Object);
            string fileName = Path.GetFullPath(Directory.GetParent(Directory.GetCurrentDirectory()).Parent.Parent + @"\test.xlsx");
            var content = File.OpenRead(fileName);
            var file = new Mock<IFormFile>();

            file.Setup(f => f.OpenReadStream()).Returns(content);
            file.Setup(f => f.FileName).Returns("test.xlsx");
            file.Setup(_ => _.Length).Returns(content.Length);

            // Act
            var result = await controller.UploadExcel(file.Object);

            Assert.IsType<OkObjectResult>(result);
        }

        [Fact]
        public async Task UploadExcel_NoFile_ReturnsBadRequest()
        {
            // Arrange
            var mockConvertExcelToObject = new Mock<IConvertExcelToObject>();
            var controller = new ExcelController(mockConvertExcelToObject.Object);
            // Act
            var result = await controller.UploadExcel(null);

            // Assert
            var badRequestResult = Assert.IsType<BadRequestObjectResult>(result);

            Assert.Equal("There was no file selected", badRequestResult.Value);
        }

        [Fact]
        public async Task UploadExcel_NoValidFile_ReturnsBadRequest()
        {
            // Arrange
            var mockConvertExcelToObject = new Mock<IConvertExcelToObject>();
            var controller = new ExcelController(mockConvertExcelToObject.Object);
            string fileName = Path.GetFullPath(Directory.GetParent(Directory.GetCurrentDirectory()).Parent.Parent + @"\test.txt");
            var content = File.OpenRead(fileName);
            var file = new Mock<IFormFile>();

            file.Setup(f => f.OpenReadStream()).Returns(content);
            file.Setup(f => f.FileName).Returns("test.txt");
            file.Setup(f => f.Length).Returns(content.Length);

            // Act
            var result = await controller.UploadExcel(file.Object);

            // Assert
            var badRequestResult = Assert.IsType<BadRequestObjectResult>(result);

            Assert.Equal("Format does not match .xlsx and .xls extensions", badRequestResult.Value);
        }
    }
}
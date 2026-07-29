using ClosedXML.Excel;
using System.IO;
using Microsoft.AspNetCore.Mvc;

namespace ExcelCrudApi.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class HomeController : ControllerBase
    {
        private readonly string _filePath = Path.Combine(Directory.GetCurrentDirectory(), "Products.xlsx");

        // Helper: Initialize file structure if missing
        private void EnsureFileExists()
        {
            if (!System.IO.File.Exists(_filePath))
            {
                using var workbook = new XLWorkbook();
                var worksheet = workbook.Worksheets.Add("Products");
                worksheet.Cell(1, 1).Value = "Id";
                worksheet.Cell(1, 2).Value = "Name";
                worksheet.Cell(1, 3).Value = "Price";
                workbook.SaveAs(_filePath);
            }
        }

        // CREATE: Add a new row
        [HttpPost]
        public IActionResult Create([FromBody] Product product)
        {
            EnsureFileExists();
            using var workbook = new XLWorkbook(_filePath);
            var worksheet = workbook.Worksheet("Products");
            
            // Find the last occupied row
            int lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 1;
            int nextRow = lastRow + 1;

            worksheet.Cell(nextRow, 1).Value = product.Id;
            worksheet.Cell(nextRow, 2).Value = product.Name;
            worksheet.Cell(nextRow, 3).Value = product.Price;

            workbook.Save();
            return Ok("Row added successfully.");
        }

        // READ: Get all items
        [HttpGet]
        public IActionResult ReadAll()
        {
            EnsureFileExists();
            var products = new List<Product>();

            using var workbook = new XLWorkbook(_filePath);
            var worksheet = workbook.Worksheet("Products");
            var rows = worksheet.RowsUsed().Skip(1); // Skip header row

            foreach (var row in rows)
            {
                products.Add(new Product
                {
                    Id = row.Cell(1).GetValue<int>(),
                    Name = row.Cell(2).GetValue<string>(),
                    Price = row.Cell(3).GetValue<decimal>()
                });
            }

            return Ok(products);
        }

        // UPDATE: Modify a row based on ID
        [HttpPut("{id}")]
        public IActionResult Update(int id, [FromBody] Product updatedProduct)
        {
            EnsureFileExists();
            using var workbook = new XLWorkbook(_filePath);
            var worksheet = workbook.Worksheet("Products");
            var rows = worksheet.RowsUsed().Skip(1);

            var rowToUpdate = rows.FirstOrDefault(r => r.Cell(1).GetValue<int>() == id);

            if (rowToUpdate == null) 
                return NotFound("Product ID not found.");

            rowToUpdate.Cell(2).Value = updatedProduct.Name;
            rowToUpdate.Cell(3).Value = updatedProduct.Price;

            workbook.Save();
            return Ok("Row updated successfully.");
        }

        // // DELETE: Remove a row by ID
        // [HttpDelete("{id}")]
        // public IActionResult Delete(int id)
        // {
        //     EnsureFileExists();
        //     using var workbook = new XLWorkbook(_filePath);
        //     var worksheet = workbook.Worksheet("Products");
        //     var rows = worksheet.RowsUsed().Skip(1);

        //     var rowToDelete = rows.FirstOrDefault(r => r.Cell(1).GetValue<int>() == id);

        //     if (rowToDelete == null) 
        //         return NotFound("Product ID not found.");

        //     // Delete the row and shift subsequent rows up
        //     rowToDelete.Delete(XLShiftDeletedCells.ShiftCellsUp);

        //     workbook.Save();
        //     return Ok("Row deleted successfully.");
        // }
    }
}

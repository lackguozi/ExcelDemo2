using ExcelDemo2.Dto;
using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Excel;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.IO;

namespace ExcelDemo2.Controllers
{
    /// <summary>
    /// shujuku
    /// </summary>
    [Route("api/[controller]/[action]")]
    [ApiController]
    public class ExcelController : ControllerBase
    {
        public ExcelController()
        {

        }

        /// <summary>
        /// 下载
        /// </summary>
        /// <returns></returns>
        [HttpPost]
        public async Task<IActionResult> TestExcel()
        {

            var list = GetItem();
            IExporter exporter = new ExcelExporter();
            var bytes =await exporter.ExportAsByteArray(list);
            var outputStream = new MemoryStream();
            await using (var stream = new MemoryStream(bytes))
            {
                using var pck = new ExcelPackage(stream);
                var ws = pck.Workbook.Worksheets[0];

                int rowCount = ws.Dimension.End.Row; // 获取行数  
                int columnCount = ws.Dimension.End.Column; // 获取列数  

                for (int column = 1; column <= columnCount; column++) // 从第一列开始遍历  
                {
                    int startMergeRow = -1;
                    string previousValue = null;

                    for (int row = 2; row <= rowCount; row++) // 从第一行开始遍历  
                    {
                        ExcelRangeBase cell = ws.Cells[row, column];
                        string currentValue = cell.Value?.ToString();

                        if (previousValue != null && previousValue == currentValue)
                        {
                            // 如果值相同，记录起始合并行  
                            if (startMergeRow == -1)
                            {
                                startMergeRow = row - 1; // 上一个单元格的行号（EPPlus 中的行号是从 1 开始的）  
                            }
                        }
                        else
                        {
                            // 如果值不同，或者当前是这一列的第一个单元格，合并之前的单元格  
                            if (startMergeRow != -1)
                            {
                                ws.Cells[startMergeRow, column, row - 1, column].Merge = true;
                                startMergeRow = -1; // 重置起始合并行  
                            }
                        }

                        previousValue = currentValue;
                    }

                    // 处理最后一组相同的单元格（如果有的话）  
                    if (startMergeRow != -1)
                    {
                        ws.Cells[startMergeRow, column, rowCount, column].Merge = true;
                    }
                }
                pck.SaveAs(outputStream);
            }
            // 将内存流的内容作为响应返回给客户端  
            outputStream.Seek(0, SeekOrigin.Begin);
            return File(outputStream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet","xiazai.xlsx");
        }
        private List<Item> GetItem()
        {
            return
            [
                 new() {
                     Code ="111",
                     Name ="测试2",
                     SkuCode = "10902",
                     SkuName="产品编制",
                     Qty = 5,
                     Sn = "8712ww2"
                 },
                 new() {
                     Code ="111",
                     Name ="测试2",
                     SkuCode = "10902",
                     SkuName="产品编制",
                     Qty = 5,
                     Sn = "8712wws"
                 },
                 new() {
                     Code ="111",
                     Name ="测试2",
                     SkuCode = "10902",
                     SkuName="产品编制",
                     Qty = 5,
                     Sn = "8712wwf"
                 },
                 new() {
                     Code ="111",
                     Name ="测试2",
                     SkuCode = "10902",
                     SkuName="产品编制",
                     Qty = 5,
                     Sn = "8712wwff"

                 },
                  new() {
                     Code ="2111",
                     Name ="测试2",
                     SkuCode = "10902",
                     SkuName="产品编制",
                     Qty = 5,
                     Sn = "8712wwff"
                 }
            ];
        }
    }
}

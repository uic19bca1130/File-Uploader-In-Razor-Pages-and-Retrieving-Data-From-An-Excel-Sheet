using ClosedXML.Excel;
using ExcelDataReader;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using System.Text;

public class IndexModel : PageModel
{
    public void OnGet()
    {
    }

    public class FileData
    {
        public string? Company { get;  set; }
        public string? city { get;  set; }
        public string? country { get;  set; }
        public string? Country { get;  set; }
        public string? isLocationIsUT { get;  set; }
        public string? isProjectInUT { get;  set; }
        public string? CollectiveAction { get;  set; }
        public string? PartnerName { get;  set; }
        public string? ContactName { get;  set; }
        public string? ProjectInfo { get;  set; }
        public string? investment { get;  set; }
        public string? focusOfProject { get;  set; }
        public string? privateProject { get;  set; }
        public string? domainName { get;  set; }
        public string? Comments { get;  set; }
    }

    public IActionResult OnPostUpload(List<IFormFile> files)
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        List<FileData> successList = new List<FileData>();
        List<FileData> errorList = new List<FileData>();

        foreach (var file in files)
        {
            if (file.Length > 0)
            {
                List<FileData> fileDataList = new List<FileData>();
                using (var stream = new MemoryStream())
                {
                    file.CopyTo(stream);
                    stream.Position = 0;

                    using (var reader = ExcelReaderFactory.CreateReader(stream, new ExcelReaderConfiguration()
                    {
                        FallbackEncoding = Encoding.GetEncoding(1252)
                    }))
                    {
                        do
                        {
                            while (reader.Read())
                            {
                                var rowData = Enumerable.Range(1, reader.FieldCount - 1)
                                    .Select(i => reader.GetValue(i)?.ToString() ?? "")
                                    .ToList();

                                var fileData = new FileData();
                                var properties = typeof(FileData).GetProperties();

                                for (int j = 0; j < properties.Length; j++)
                                {
                                    if (j < rowData.Count)
                                    {
                                        properties[j].SetValue(fileData, rowData[j]);
                                    }
                                }

                                if (string.IsNullOrEmpty(fileData.ProjectInfo) || string.IsNullOrEmpty(fileData.city) || string.IsNullOrEmpty(fileData.Country))
                                {
                                    errorList.Add(fileData);
                                }
                                else
                                {
                                    successList.Add(fileData);
                                }

                                fileDataList.Add(fileData);
                            }
                        } while (reader.NextResult());
                    }
                }

                if (fileDataList.Count > 0)
                {
                    var response = new Response
                    {
                        SuccessCount = successList.Count,
                        FailureCount = errorList.Count,
                        ExcelFile = GenerateExcelFile(errorList)
                    };

                    return new JsonResult(response);
                }
            }
        }

        return new OkResult();
    }

    private byte[] GenerateExcelFile(List<FileData> dataList)
    {
        using (var workbook = new XLWorkbook())
        {
            var worksheet = workbook.Worksheets.Add("DataList");
            var headerRow = typeof(FileData).GetProperties().Select(prop => prop.Name).ToList();

            for (int i = 0; i < headerRow.Count; i++)
            {
                worksheet.Cell(1, i + 1).Value = headerRow[i];
            }

            for (int i = 0; i < dataList.Count; i++)
            {
                var rowData = typeof(FileData).GetProperties().Select(prop =>
                {
                    var value = prop.GetValue(dataList[i]);
                    return value?.ToString() ?? "no data";
                }).ToList();

                for (int j = 0; j < rowData.Count; j++)
                {
                    worksheet.Cell(i + 2, j + 1).Value = rowData[j];
                }
            }

            worksheet.Columns().AdjustToContents();

            using (var stream = new MemoryStream())
            {
                workbook.SaveAs(stream);
                return stream.ToArray();
            }
        }
    }

    public class Response
    {
        public int SuccessCount { get; set; }
        public int FailureCount { get; set; }
        public byte[] ExcelFile { get; set; }
    }
}

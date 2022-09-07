
 using ClosedXML.Excel;
using ClosedXmlExample.Models;

List<Student> _oStudents = new List<Student>();


   for (int i = 1; i <= 9; i++)
   {
        _oStudents.Add(new Student()
        {
           Id = i,
           Name = "Student" + i,
           Roll = "100" + i,
           Address="test"+i
        });
    }



using (var workbook = new XLWorkbook())
{
    var worksheet = workbook.Worksheets.Add("Student");
    var currentRow = 1;

    worksheet.Cell(currentRow, 1).Value = "StudentId";
    worksheet.Cell(currentRow, 2).Value = "Name";
    worksheet.Cell(currentRow, 3).Value = "Roll";
    worksheet.Cell(currentRow, 4).Value = "Address";

    foreach (var Student in _oStudents)
    {
        currentRow++;
        worksheet.Cell(currentRow, 1).Value = Student.Id;
        worksheet.Cell(currentRow, 2).Value = Student.Name;
        worksheet.Cell(currentRow, 3).Value = Student.Roll;
        worksheet.Cell(currentRow, 4).Value = Student.Address;
    }

    using (var stream = new MemoryStream())
    {
        workbook.SaveAs(stream);
        File.WriteAllBytes("C:\\Users\\User\\Desktop\\test\\excel.xlsx", StreamHelpers.GetBytes(stream));
    }
}
        Console.WriteLine("Hello, World!");

using System;
using System.IO;
using System.Text;
using OfficeOpenXml;

class Program
{
    static void Main()
    {
        
        string directoryPath = "/Users/duypc/Downloads/Data";
        string searchPattern = "*.xlsx";

        try
        {
            string[] files = Directory.GetFiles(directoryPath, searchPattern);
            var test = new GenerateExcelToSql();
            StringBuilder sqlString = new StringBuilder();
            foreach (string file in files)
            {
                sqlString.Append(test.GenerateSqlFromExcel(file));
            }
            test.CheckExistFile(sqlString.ToString());
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Lỗi: {ex.Message}");
        }
    }
}

public class GenerateExcelToSql
{
    private string exportTxt = "/Users/duypc/Desktop/sql.txt"; 
    
    public string GenerateSqlFromExcel(string file)
    {
        
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        
        string fileName = Path.GetFileNameWithoutExtension(file);

        
        using (var excelPackage = new ExcelPackage(new FileInfo(file)))
        {
            
            var worksheet = excelPackage.Workbook.Worksheets[0];

            
            StringBuilder sqlBuilder = new StringBuilder();
            sqlBuilder.Append($"INSERT INTO \"Catalog\".\"{fileName}\" (");

            
            for (var col = 2; col <= worksheet.Dimension.End.Column; col++)
            {
                var field = worksheet.Cells[1, col].Text;
                sqlBuilder.Append(col == worksheet.Dimension.End.Column ? $"\"{field}\") \n" : $"\"{field}\", ");
            }

            
            sqlBuilder.AppendLine("VALUES");

            
            for (var row = 2; row <= worksheet.Dimension.End.Row; row++)
            {
                sqlBuilder.Append("(");
                for (var col = 2; col <= worksheet.Dimension.End.Column; col++)
                {
                    var value = worksheet.Cells[row, col].Value; 
                    if (value is string && !value.ToString().Equals("null"))
                    {
                        sqlBuilder.Append(col == worksheet.Dimension.End.Column ? $"'{value}') \n" : $"'{value}', ");
                    }
                    else
                    {
                        sqlBuilder.Append(col == worksheet.Dimension.End.Column ? $"{value}) \n" : $"{value}, ");
                    }
                }
                if (row != worksheet.Dimension.End.Row)
                {
                    sqlBuilder.AppendLine(",");
                }
            }

            return sqlBuilder.ToString();
        }
    }

    public void CheckExistFile(string sqlBuilder)
    {
        if (!File.Exists(exportTxt))
        {
            try
            {
                    
                using (FileStream fs = File.Create(exportTxt))
                {
                    Console.WriteLine($"File {exportTxt} has been created.");
                }
            }
            catch (Exception e)
            {
                Console.WriteLine($"An error occurred while creating the file: {e.Message}");
            }
        }
        else
        {
            Console.WriteLine($"File {exportTxt} already exists.");
        }

            
        try
        {
            using (StreamWriter writer = new StreamWriter(exportTxt, append: true))
            {
                writer.Write(sqlBuilder);
            }

            Console.WriteLine($"SQL statements successfully written to {exportTxt}");
        }
        catch (Exception e)
        {
            Console.WriteLine($"Error writing to file: {e.Message}");
        }
    }
}

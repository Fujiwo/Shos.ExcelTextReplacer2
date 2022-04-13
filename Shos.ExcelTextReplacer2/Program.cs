#nullable enable

using Excel = Microsoft.Office.Interop.Excel;

namespace Shos.ExcelTextReplacer2
{
    static class Helper
    {
        public static bool IsBetween(this int @this, int minimum, int maximum) => @this >= minimum && @this <= maximum;

        public static string? ToFullPath(this string filePath)
            => File.Exists(filePath) ? Path.GetFullPath(filePath) : null;
    }

    record struct Parameters(string excelFilePath, int idColumn, int column);

    static class Program
    {
        static void Main(string[] args)
        {
            var parameters = GetParameters(args);

            if (parameters is null)
                Usage();
            else
                Replace(parameters.Value.targetParameters, parameters.Value.inputParameters);
        }

        static Parameters? GetParameters(string columnsText)
        {
            var columnTexts = columnsText.Split(',');
            if (columnTexts.Length >= 3) {
                var filePath = columnTexts[0].ToFullPath();
                if (!string.IsNullOrWhiteSpace(filePath) && int.TryParse(columnTexts[1], out var column1) && int.TryParse(columnTexts[2], out var column2))
                    return new Parameters(filePath, column1, column2);
            }
            return null;
        }

        static (Parameters targetParameters, Parameters inputParameters)? GetParameters(string[] args)
        {
            Parameters? targetParameters = null;
            Parameters? inputParameters  = null;

            for (var index = 0; index < args.Length; index++) {
                switch (args[index]) {
                    case "-t":
                    case "-T":
                    case "/t":
                    case "/T":
                        if (args.Length > index + 1) {
                            targetParameters = GetParameters(args[index + 1]);
                            index++;
                        }
                        break;

                    case "-i":
                    case "-I":
                    case "/i":
                    case "/I":
                        if (args.Length > index + 1) {
                            inputParameters = GetParameters(args[index + 1]);
                            index++;
                        }
                        break;
                }
            }
            if (targetParameters is null || inputParameters is null)
                return null;
            return (targetParameters.Value, inputParameters.Value);
        }

        static void Replace(Parameters targetParameters, Parameters inputParameters)
        {
            var excel = new Excel.Application();
            excel.Visible = true;
            Replace(excel, targetParameters, inputParameters);
            excel.Quit();
        }

        static void Replace(Excel.Application excel, Parameters targetParameters, Parameters inputParameters)
        {
            var targetWorkbook = excel.Workbooks.Open(targetParameters.excelFilePath);
            if (targetWorkbook is null)
                return;

            var inputWorkbook  = excel.Workbooks.Open(inputParameters .excelFilePath);
            if (inputWorkbook is null) {
                targetWorkbook.Close(true);
                return;
            }

            Replace((targetWorkbook, targetParameters.idColumn, targetParameters.column), (inputWorkbook, inputParameters.idColumn, inputParameters.column));

            inputWorkbook .Close(true);
            targetWorkbook.Close(true);
        }

        static void Replace((Excel.Workbook workbook, int idColumn, int column) target, (Excel.Workbook workbook, int idColumn, int column) input)
        {
            var inputData = GetInputData(input);
            foreach (Excel.Worksheet sheet in target.workbook.Sheets)
                Replace(inputData, sheet, target.idColumn, target.column);
        }

        static Dictionary<string, string> GetInputData((Excel.Workbook workbook, int idColumn, int column) input)
        {
            Dictionary<string, string> inputData = new();

            foreach (Excel.Worksheet sheet in input.workbook.Sheets) {
                var columnCount = sheet.UsedRange.Columns.Count;

                if (!input.idColumn.IsBetween(1, columnCount) || !input.column.IsBetween(1, columnCount))
                    continue;

                var rowCount = sheet.UsedRange.Rows.Count;

                for (var row = 1; row <= rowCount; row++) {
                    var id        = ToText(sheet.Cells[row, input.idColumn]);
                    var value     = ToText(sheet.Cells[row, input.column  ]);
                    inputData[id] = value;
                }
            }
            return inputData;
        }

        static void Replace(Dictionary<string, string> inputData, Excel.Worksheet sheet, int idColumn, int column)
        {
            var columnCount = sheet.UsedRange.Columns.Count;
            if (!idColumn.IsBetween(1, columnCount) || !column.IsBetween(1, columnCount))
                return;

            var rowCount = sheet.UsedRange.Rows.Count;

            for (var row = 1; row <= rowCount; row++) {
                var id = ToText(sheet.Cells[row, idColumn]);
                if (inputData.TryGetValue(id, out string value)) {
                    var range = sheet.Cells[row, column] as Excel.Range;
                    if (range is not null)
                        range.Value = value;
                }
            }
        }

        static string ToText(object cell)
        {
            var range = cell as Excel.Range;
            if (range is not null) {
                dynamic value = range.Value;
                var text = Convert.ToString(value);
                if (!string.IsNullOrWhiteSpace(text))
                    return text;
            }
            return "";
        }

        static void Usage() => Console.WriteLine(
            "Usage:\nShos.ExcelTextReplacer2 -t targetExcelFilePath,targetIdColumn,targetColumn -i inputExcelFilePath,inputIdColumn,inputColumn\n" +
            "\n" +
            "-t targetExcelFilePath,targetIdColumn,targetColumn\n\tTarget Excel file path,Target id column index,Target column index.\n" +
            "-i inputExcelFilePath,inputIdColumn,inputColumn\n\tInput Excel file path,Input id column index,Input column index.\n" +
            "\n" +
            "ex.\n" +
            "\n" +
            "Shos.ExcelTextReplacer2 -t target.xlsx,0,1 -i input.xlsx,2,3"
        );
    }
}

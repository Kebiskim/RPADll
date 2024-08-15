using System;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

class Program
{
    static void Main(string[] args)
    {
        Excel.Application excelApp = null;

        try
        {
            // Excel 애플리케이션 시작
            excelApp = new Excel.Application();
            excelApp.Visible = true; // Excel UI를 표시

            // 파일 열기
            string filePath = @"C:\LOCAL_RPA\MacroTest.xlsx";
            Excel.Workbook workbook = excelApp.Workbooks.Open(filePath);

            // 첫 번째 워크시트의 이름을 출력하여 내용 확인
            if (workbook.Worksheets.Count > 0)
            {
                Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Worksheets[1];
                Console.WriteLine($"첫 번째 워크시트 이름: {worksheet.Name}");
            }
            else
            {
                Console.WriteLine("워크북에 시트가 없습니다.");
            }

            // 작업 후 Excel 프로세스 종료
            workbook.Close(false);
            Marshal.ReleaseComObject(workbook);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"오류 발생: {ex.Message}");
        }
        finally
        {
            // Excel 애플리케이션 종료
            if (excelApp != null)
            {
                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);
            }

            // Console 창이 즉시 닫히지 않도록 합니다.
            Console.WriteLine("Press any key to exit...");
            Console.ReadLine();
        }
    }
}




/*
엑셀 Office.dll 파일이 존재하지 않는 경우(로컬에 Office가 설치되지 않은 경우)
에도 대응할 수 있도록 EPPlus를 사용하자.
하지만 EPPlus는 Excel을 백그라운드에서만 열 수 있다.
*/

// using OfficeOpenXml;
// using System;
// using System.IO;

// class Program
// {
//     static void Main(string[] args)
//     {
//         // Excel 패키지 사용 전 권한 부여
//         ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

//         // Excel 파일 경로 설정
//         string filePath = @"C:\LOCAL_RPA\MacroTest.xlsx";
//         FileInfo fileInfo = new FileInfo(filePath);

//         // 파일이 실제로 존재하는지 확인
//         if (!fileInfo.Exists)
//         {
//             Console.WriteLine($"파일이 존재하지 않습니다: {filePath}");
//             return;
//         }

//         try
//         {
//             using (ExcelPackage package = new ExcelPackage(fileInfo))
//             {
//                 ExcelWorkbook workbook = package.Workbook;

//                 if (workbook != null && workbook.Worksheets.Count > 0)
//                 {
//                     Console.WriteLine("Excel 파일이 정상적으로 열렸습니다.");

//                     // 첫 번째 워크시트의 이름을 출력하여 내용 확인
//                     var worksheet = workbook.Worksheets[0];
//                     Console.WriteLine($"첫 번째 워크시트 이름: {worksheet.Name}");
//                 }
//                 else
//                 {
//                     Console.WriteLine("워크북을 여는 데 실패했습니다.");
//                 }
//             }
//         }
//         catch (Exception ex)
//         {
//             Console.WriteLine($"파일을 여는 도중 오류 발생: {ex.Message}");
//         }
//     }
// }

using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;
using System.Activities;
//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;
//using System.Activities;
using System.ComponentModel;
//using System.Data;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.IO;
//using Microsoft.CSharp.RuntimeBinder;

//using System.Runtime.InteropServices;
namespace CustomExcelOperations
{
    public class FormatPainter : CodeActivity
    {
        [Category("Source")]
        [RequiredArgument]
        public InArgument<string> SourceFile{ get; set; }

        [RequiredArgument]
        [Category("Source")]
        public InArgument<string> SourceSheet { get; set; }

        [Category("Destination")]
        [RequiredArgument]
        public InArgument<string> DestinationFile{ get; set; }

        [RequiredArgument]
        [Category("Destination")]
        public InArgument<string> DestinationSheet { get; set; }


        protected override void Execute(CodeActivityContext context)
        {
            
            Application excelApp = new Application();
            excelApp.DisplayAlerts = false;
            
            var source_file = SourceFile.Get(context);
            var dest_file = DestinationFile.Get(context);
            //UiPath.Excel.Activities.ExcelGetTableRange
            var Source_SheetName = SourceSheet.Get(context);
            var Destination_SheetName = DestinationSheet.Get(context);
            Workbook Sourcefile_Workbook = excelApp.Workbooks.Add(System.Reflection.Missing.Value);
            Workbook Destination_Workbook = excelApp.Workbooks.Add(System.Reflection.Missing.Value);

            
            var flagToCheckSourceAssignment = true;
  
            try
            {
                Application application = null;
                application = (Application)CustomMarshal.GetActiveObject("Excel.Application");
                var WorkbookList = application.Workbooks;
                var fileShortName = Path.GetFileName(source_file);
                
                foreach (Workbook book in WorkbookList)
                {
                    //Console.WriteLine(book.Name + "\n");
                    if (book.Name.Equals(fileShortName))
                    {
                        Sourcefile_Workbook = book;
                        flagToCheckSourceAssignment = false;
                    }

                }

            }
            catch(Exception e)
            {
                Console.WriteLine(e.Message);
            }

            if (flagToCheckSourceAssignment)
            {
                try
                {
                    Sourcefile_Workbook = excelApp.Workbooks.Open(source_file,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                }
            }


            var flagToCheckAssignment = true;
            try
            {
                Application application = null;
                application = (Application)CustomMarshal.GetActiveObject("Excel.Application");
                application.DisplayAlerts = false;
                var WorkbookList = application.Workbooks;
                //var WorkbookList = MSExcelWorkbookRunningInstances.Enum();
                var fileShortName = Path.GetFileName(dest_file);
                foreach (Workbook book in WorkbookList)
                {
                    //Console.WriteLine(book.FullName + "\n");
                    if (book.Name.Equals(fileShortName))
                    {
                        Destination_Workbook = book;
                        flagToCheckAssignment = false;
                    }
                   
                }
            }
            catch(Exception e)
            {
                Console.WriteLine(e.Message);
            }
            if (flagToCheckAssignment)
            {
                try
                {
                    Destination_Workbook = excelApp.Workbooks.Open(dest_file,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                }
                catch(Exception e)
                {
                    Console.WriteLine(e.Message);
                }
            }

            object misValue = System.Reflection.Missing.Value;
            Worksheet Sourceworksheet = (Worksheet)Sourcefile_Workbook.Worksheets[Source_SheetName];
            Range range = Sourceworksheet.UsedRange;
            range.Copy();
            
            
            Worksheet Destinationworksheet = (Worksheet)Destination_Workbook.Worksheets[Destination_SheetName];
            Range DestinationRange = Destinationworksheet.UsedRange;
            DestinationRange.PasteSpecial(XlPasteType.xlPasteFormats);

           
            if (flagToCheckSourceAssignment)
            {
                Sourcefile_Workbook.Close(false, misValue, misValue);
            }

            if (flagToCheckAssignment)
            {
                Destination_Workbook.Close(true, misValue, misValue);
            }
            excelApp.Quit();
        }
    }
}

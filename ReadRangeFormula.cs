using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Activities;
using System.ComponentModel;
using System.Data;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace CustomExcelOperations
{
    public class ReadRangeFormula : CodeActivity
    {

        [Category("Input")]
        [RequiredArgument]
        public InArgument<string> FilePath { get; set; }

        [RequiredArgument]
        [Category("Input")]
        public InArgument<string> SheetName { get; set; }

        //[Category("Input")]
        //public InArgument<double> OpsRange { get; set; }

       

        [Category("Output")]
        public OutArgument<System.Object[,]> Result { get; set; }


        protected override void Execute(CodeActivityContext context)
        //public void ExecuteOB()
        {
            Application excelApp = new Application();
            var file_Path = FilePath.Get(context);
            var sheet_Name = SheetName.Get(context);
            Workbook workbook = excelApp.Workbooks.Open(file_Path,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            // The key line:
            Worksheet worksheet = (Worksheet)workbook.Worksheets[sheet_Name];
            Range range = worksheet.UsedRange;
            System.Object[,] OutputData = (object[,])range.Formula;
            var typeOfVar = OutputData.GetType();
            Result.Set(context, OutputData);


            
        }


    }
}

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

        [Category("Output")]
        public OutArgument<System.String> Log { get; set; }


        protected override void Execute(CodeActivityContext context)
        //public void ExecuteOB()
        {
            Application excelApp = new Application();
            excelApp.Visible = false;
            excelApp.UserControl = false;
            var file_Path = FilePath.Get(context);
            var sheet_Name = SheetName.Get(context);
            Workbook workbook = excelApp.Workbooks.Open(file_Path,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing);

          
            
            // The key line:
            Worksheet worksheet = (Worksheet)workbook.Worksheets[sheet_Name];
            //workbook.UserControl = false;
            object misValue = System.Reflection.Missing.Value;

            Range range = worksheet.UsedRange;
            
            try
            {
                System.Object[,] OutputData = (object[,])range.Formula;
                //var typeOfVar = OutputData.GetType();
                Result.Set(context, OutputData);
                Log.Set(context,sheet_Name + " sheet has been read");
                
            }
            catch
            {
                System.Object[,] OutputData = new object[0, 0];
                Result.Set(context, OutputData);
                Log.Set(context, "Unable to read sheet: " + sheet_Name + " ");
                Console.WriteLine("Unable to read sheet: " + sheet_Name + " ");

            }
            finally
            {
                
                workbook.Close(false, misValue, misValue);
            }
            

            
        }


    }
}

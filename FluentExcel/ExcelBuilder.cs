using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using Microsoft.Office.Interop.Excel;

namespace FluentExcel
{
    public class ExcelBuilder
    {
        private Application application;
        private Workbook lastWorkbook;
        private Worksheet lastWorkSheet;

        ExcelBuilder() { }

        public static ExcelBuilder Begin()
        {
            var builder = new ExcelBuilder
            {
                application = new Application()
            };
            return builder;
        }

        public ExcelBuilder AddWorkbook(XlWBATemplate xlWBATemplate = XlWBATemplate.xlWBATWorksheet)
        {
            lastWorkbook = application.Workbooks.Add(xlWBATemplate);
            return this;
        }

        public ExcelBuilder AddWorkSheet<T>(IEnumerable<T> items, params Expression<Func<T, object>>[] expressions)
        {
            if (lastWorkSheet == null)
            {
                lastWorkSheet = (Worksheet)lastWorkbook.Worksheets[1];
            }
            else
            {
                lastWorkSheet = lastWorkbook.Worksheets.Add(new Worksheet());
            }
            var columnNames = new List<string>();
            foreach (var expression in expressions)
            {
                columnNames.Add(GetMemberName(expression));
            }
            var columnCounter = 1;
            var rowCounter = 1;

            foreach (var columnName in columnNames)
            {
                lastWorkSheet.Cells[rowCounter, columnCounter] = columnName;
                foreach (var item in items)
                {
                    rowCounter++;
                    lastWorkSheet.Cells[rowCounter, columnCounter] = typeof(T).GetProperty(columnName).GetValue(item).ToString();
                }
                rowCounter = 1;
                columnCounter++;
            }
            return this;
        }

        public ExcelBuilder SaveWorkbook(string filePathName)
        {
            lastWorkbook.SaveAs(filePathName);
            lastWorkbook.Close();
            lastWorkbook = null;
            lastWorkSheet = null;
            return this;
        }

        public void End()
        {
            application.Quit();
            application = null;
        }

        internal static string GetMemberName(LambdaExpression lambdaExpression)
        {
            string result;

            MemberExpression memberExpression;
            if (lambdaExpression.Body is UnaryExpression)
            {
                var unaryExpression = (UnaryExpression)lambdaExpression.Body;
                memberExpression = (MemberExpression)unaryExpression.Operand;
            }
            else
            {
                memberExpression = (MemberExpression)lambdaExpression.Body;
            }

            result = memberExpression.Member.Name;

            return result;
        }
    }
}

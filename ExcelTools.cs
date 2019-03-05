using System;
using System.Data;
using ExcelTools.Drivers;

namespace ExcelTools
{
    public class ExcelTool
    {
        protected string filepath;

        protected IExcelTools driver;

        public ExcelTool(string filepath) {
            this.filepath = filepath;
        }

        public void FromDataTable(DataTable data) {
            driver.FromDataTable(data);
        }

        public DataTable ToDataTable() {
            return driver.ToDataTable();
        }

        public void SetDriver(string driverName) {
            IExcelTools driverObject;
            switch (driverName) {

                case "EPPlus":
                    throw new Exception("Not implemented yet");

                case "NPOI":
                default:
                    driverObject = new NPOIDriver(filepath);
                    break;
            }
            driver = driverObject;
        }

    }
}

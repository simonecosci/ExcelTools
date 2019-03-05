using System.Data;

namespace ExcelTools
{
    public interface IExcelTools
    {
        void FromDataTable(DataTable data);
        DataTable ToDataTable();
    }
}

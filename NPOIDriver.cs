using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Data;
using System.IO;

namespace ExcelTools.Drivers
{
    public class NPOIDriver : ExcelTool, IExcelTools
    {
        protected new IExcelTools driver = null; 

        public NPOIDriver(string filepath) : base(filepath) { }

        public new void FromDataTable(DataTable data)
        {
            try
            {
                if (data != null && data.Rows.Count > 0)
                {
                    IWorkbook workbook = null;
                    ISheet worksheet = null;

                    using (FileStream stream = new FileStream(filepath, FileMode.Create, FileAccess.ReadWrite))
                    {
                        string Ext = Path.GetExtension(filepath);
                        switch (Ext.ToLower())
                        {
                            case ".xls":
                                HSSFWorkbook workbookH = new HSSFWorkbook();
                                NPOI.HPSF.DocumentSummaryInformation dsi = NPOI.HPSF.PropertySetFactory.CreateDocumentSummaryInformation();
                                dsi.Company = "Pierotucci"; dsi.Manager = "Divisione Informatica";
                                workbookH.DocumentSummaryInformation = dsi;
                                workbook = workbookH;
                                break;

                            case ".xlsx": workbook = new XSSFWorkbook(); break;
                        }

                        worksheet = workbook.CreateSheet(data.TableName);

                        // columns titles
                        int iRow = 0;
                        if (data.Columns.Count > 0)
                        {
                            int iCol = 0;
                            IRow row = worksheet.CreateRow(iRow);
                            foreach (DataColumn columna in data.Columns)
                            {
                                ICell cell = row.CreateCell(iCol, CellType.String);
                                cell.SetCellValue(columna.ColumnName);
                                iCol++;
                            }
                            iRow++;
                        }

                        ICellStyle _doubleCellStyle = workbook.CreateCellStyle();
                        _doubleCellStyle.DataFormat = workbook.CreateDataFormat().GetFormat("#0.###");

                        ICellStyle _intCellStyle = workbook.CreateCellStyle();
                        _intCellStyle.DataFormat = workbook.CreateDataFormat().GetFormat("#0");

                        ICellStyle _boolCellStyle = workbook.CreateCellStyle();
                        _boolCellStyle.DataFormat = workbook.CreateDataFormat().GetFormat("BOOLEAN");

                        ICellStyle _dateCellStyle = workbook.CreateCellStyle();
                        _dateCellStyle.DataFormat = workbook.CreateDataFormat().GetFormat("dd-MM-yyyy");

                        ICellStyle _dateTimeCellStyle = workbook.CreateCellStyle();
                        _dateTimeCellStyle.DataFormat = workbook.CreateDataFormat().GetFormat("dd-MM-yyyy HH:mm:ss");

                        foreach (DataRow row in data.Rows)
                        {
                            IRow line = worksheet.CreateRow(iRow);
                            int iCol = 0;
                            foreach (DataColumn column in data.Columns)
                            {
                                ICell cell = null;
                                object cellValue = row[iCol];
                                switch (column.DataType.ToString())
                                {
                                    case "System.Boolean":
                                        if (cellValue != DBNull.Value)
                                        {
                                            cell = line.CreateCell(iCol, CellType.Boolean);

                                            if (Convert.ToBoolean(cellValue)) { cell.SetCellFormula("TRUE()"); }
                                            else { cell.SetCellFormula("FALSE()"); }

                                            cell.CellStyle = _boolCellStyle;
                                        }
                                        break;
                                    case "System.Int32":
                                        if (cellValue != DBNull.Value)
                                        {
                                            cell = line.CreateCell(iCol, CellType.Numeric);
                                            cell.SetCellValue(Convert.ToInt32(cellValue));
                                            cell.CellStyle = _intCellStyle;
                                        }
                                        break;
                                    case "System.Int64":
                                        if (cellValue != DBNull.Value)
                                        {
                                            cell = line.CreateCell(iCol, CellType.Numeric);
                                            cell.SetCellValue(Convert.ToInt64(cellValue));
                                            cell.CellStyle = _intCellStyle;
                                        }
                                        break;
                                    case "System.Decimal":
                                        if (cellValue != DBNull.Value)
                                        {
                                            cell = line.CreateCell(iCol, CellType.Numeric);
                                            cell.SetCellValue(Convert.ToDouble(cellValue));
                                            cell.CellStyle = _doubleCellStyle;
                                        }
                                        break;
                                    case "System.Double":
                                        if (cellValue != DBNull.Value)
                                        {
                                            cell = line.CreateCell(iCol, CellType.Numeric);
                                            cell.SetCellValue(Convert.ToDouble(cellValue));
                                            cell.CellStyle = _doubleCellStyle;
                                        }
                                        break;
                                    case "System.DateTime":
                                        if (cellValue != DBNull.Value)
                                        {
                                            cell = line.CreateCell(iCol, CellType.Numeric);
                                            cell.SetCellValue(Convert.ToDateTime(cellValue));
                                            DateTime cDate = Convert.ToDateTime(cellValue);
                                            if (cDate != null && cDate.Hour > 0) { cell.CellStyle = _dateTimeCellStyle; }
                                            else { cell.CellStyle = _dateCellStyle; }
                                        }
                                        break;
                                    default:
                                        if (cellValue != DBNull.Value)
                                        {
                                            cell = line.CreateCell(iCol, CellType.String);
                                            cell.SetCellValue(Convert.ToString(cellValue));
                                        }
                                        break;
                                }
                                iCol++;
                            }
                            iRow++;
                        }
                        workbook.Write(stream);
                        stream.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public new DataTable ToDataTable()
        {
            DataTable table = null;
            try
            {
                if (File.Exists(filepath))
                {

                    IWorkbook workbook = null;
                    ISheet worksheet = null;
                    string first_sheet_name = "";

                    using (FileStream FS = new FileStream(filepath, FileMode.Open, FileAccess.Read))
                    {
                        workbook = WorkbookFactory.Create(FS);
                        worksheet = workbook.GetSheetAt(0);
                        first_sheet_name = worksheet.SheetName;

                        table = new DataTable(first_sheet_name);
                        table.Rows.Clear();
                        table.Columns.Clear();

                        for (int rowIndex = 0; rowIndex <= worksheet.LastRowNum; rowIndex++)
                        {
                            DataRow NewReg = null;
                            IRow row = worksheet.GetRow(rowIndex);
                            IRow row2 = null;
                            IRow row3 = null;

                            if (rowIndex == 0)
                            {
                                row2 = worksheet.GetRow(rowIndex + 1);
                                row3 = worksheet.GetRow(rowIndex + 2);
                            }

                            if (row != null) //null is when the row only contains empty cells 
                            {
                                if (rowIndex > 0) NewReg = table.NewRow();

                                int colIndex = 0;
                                foreach (ICell cell in row.Cells)
                                {
                                    object valorCell = null;
                                    string cellType = "";
                                    string[] cellType2 = new string[2];

                                    if (rowIndex == 0)
                                    {
                                        for (int i = 0; i < 2; i++)
                                        {
                                            ICell cell2 = null;
                                            if (i == 0) { cell2 = row2.GetCell(cell.ColumnIndex); }
                                            else {
                                                if (row3 != null)
                                                    cell2 = row3.GetCell(cell.ColumnIndex);
                                            }

                                            if (cell2 == null)
                                            {
                                                continue;
                                            }
                                            switch (cell2.CellType)
                                            {
                                                case CellType.Blank: break;
                                                case CellType.Boolean: cellType2[i] = "System.Boolean"; break;
                                                case CellType.String: cellType2[i] = "System.String"; break;
                                                case CellType.Numeric:
                                                    if (HSSFDateUtil.IsCellDateFormatted(cell2)) { cellType2[i] = "System.DateTime"; }
                                                    else
                                                    {
                                                        cellType2[i] = "System.Double";
                                                    }
                                                    break;

                                                case CellType.Formula:
                                                    bool continuar = true;
                                                    switch (cell2.CachedFormulaResultType)
                                                    {
                                                        case CellType.Boolean: cellType2[i] = "System.Boolean"; break;
                                                        case CellType.String: cellType2[i] = "System.String"; break;
                                                        case CellType.Numeric:
                                                            if (HSSFDateUtil.IsCellDateFormatted(cell2)) { cellType2[i] = "System.DateTime"; }
                                                            else
                                                            {
                                                                try
                                                                {
                                                                    if (cell2.CellFormula == "TRUE()") { cellType2[i] = "System.Boolean"; continuar = false; }
                                                                    if (continuar && cell2.CellFormula == "FALSE()") { cellType2[i] = "System.Boolean"; continuar = false; }
                                                                    if (continuar) { cellType2[i] = "System.Double"; continuar = false; }
                                                                }
                                                                catch { }
                                                            }
                                                            break;
                                                    }
                                                    break;
                                                default:
                                                    cellType2[i] = "System.String"; break;
                                            }
                                        }

                                        if (cellType2[0] == cellType2[1]) { cellType = cellType2[0]; }
                                        else
                                        {
                                            if (cellType2[0] == null) cellType = cellType2[1];
                                            if (cellType2[1] == null) cellType = cellType2[0];
                                            if (cellType == "") cellType = "System.String";
                                        }
                                        if (cellType == null)
                                        {
                                            cellType = "System.String";
                                        }
                                        string colName = "Column_{0}";
                                        try { colName = cell.StringCellValue; }
                                        catch { colName = string.Format(colName, colIndex); }

                                        foreach (DataColumn col in table.Columns)
                                        {
                                            if (col.ColumnName == colName)
                                            {
                                                colName = string.Format("{0}_{1}", colName, colIndex);
                                                break;
                                            }
                                        }

                                        DataColumn code = new DataColumn(colName, Type.GetType(cellType));
                                        table.Columns.Add(code);
                                        colIndex++;
                                    }
                                    else
                                    {
                                        switch (cell.CellType)
                                        {
                                            case CellType.Blank: valorCell = DBNull.Value; break;
                                            case CellType.Boolean: valorCell = cell.BooleanCellValue; break;
                                            case CellType.String: valorCell = cell.StringCellValue; break;
                                            case CellType.Numeric:
                                                if (HSSFDateUtil.IsCellDateFormatted(cell)) { valorCell = cell.DateCellValue; }
                                                else { valorCell = cell.NumericCellValue; }
                                                break;
                                            case CellType.Formula:
                                                switch (cell.CachedFormulaResultType)
                                                {
                                                    case CellType.Blank: valorCell = DBNull.Value; break;
                                                    case CellType.String: valorCell = cell.StringCellValue; break;
                                                    case CellType.Boolean: valorCell = cell.BooleanCellValue; break;
                                                    case CellType.Numeric:
                                                        if (HSSFDateUtil.IsCellDateFormatted(cell)) { valorCell = cell.DateCellValue; }
                                                        else { valorCell = cell.NumericCellValue; }
                                                        break;
                                                }
                                                break;
                                            default: valorCell = cell.StringCellValue; break;
                                        }
                                        if (cell.ColumnIndex <= table.Columns.Count - 1)
                                        {
                                            NewReg[cell.ColumnIndex] = valorCell;
                                        }
                                    }
                                }
                            }
                            if (rowIndex > 0 && NewReg != null)
                                table.Rows.Add(NewReg);
                        }
                        table.AcceptChanges();
                    }
                }
                else
                {
                    throw new Exception("ERROR 404: File not found.");
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return table;
        }
    }
}

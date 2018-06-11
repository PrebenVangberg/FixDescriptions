using System;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

public class Excel
{

    string path;
    _Application excel = new _Excel.Application();
    Workbook wb;
    Worksheet ws;

    public Excel(string path)
	{
        this.path = path;
        wb = excel.Workbooks.Open(path);
        SaveAs(AppDomain.CurrentDomain.BaseDirectory + "out.xlsx");
        wb.Close();
        wb = excel.Workbooks.Open(AppDomain.CurrentDomain.BaseDirectory + "out.xlsx");
        ws = wb.Worksheets[1];
	}

    public string ReadCell(int x, int y)
    {
        x++;
        y++;
        if(ws.Cells[x,y].Value2 != null)
        return ws.Cells[x, y].Value2;
        return "";
    }

    public void WriteCell(int x, int y, String s)
    {
        x++;
        y++;
        ws.Cells[x, y].Value2 = s;        
    }

    public void Save()
    {
        wb.Save();
    }

    public int getColumns()
    {
        return ws.Columns.Count;
    }

    public int getRows()
    {
        return ws.Rows.Count;
    }

    public void Close()
    {
        wb.Close(true);
        excel.Quit();
    }

    public void SaveAs(string path)
    {
        wb.SaveAs(path);
    }
}

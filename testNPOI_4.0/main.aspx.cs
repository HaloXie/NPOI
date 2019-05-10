using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Text;
using System.Data;
using System.IO;

//2003Excel
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;

//用于Excel 转换成 Html
using NPOI.SS.Converter;

//用于读取Excel中的图片 
using NPOI.POIFS.FileSystem;


//这个主要是针对 I***接口来实现的 因为在4.0中淡化了对象，采用接口的形式比较多
using NPOI.SS;
using NPOI.SS.UserModel;
using NPOI.SS.Util;

//主要是针对Excel中的版本信息或者是作者之类的
using NPOI.HPSF;

//ZIP
using NPOI.OpenXml4Net.OPC;
using System.Net.Mime;

//用于设置一些自定义属性
//using NPOI.HPSF;
//using NPOI.POIFS.FileSystem;

//Excel 2007
using NPOI.XSSF.UserModel;

//word 2007
using NPOI.XWPF.UserModel;
using NPOI.XWPF.Extractor;

//用于读取
//using NPOI.OpenXml4Net.OPC;
using NPOI.OpenXmlFormats.Wordprocessing;

///谢明昊 收集整理

namespace testNPOI_4._0
{
    public partial class main : System.Web.UI.Page
    {

        //注意：操作Excel2003与操作Excel2007使用的是不同的命名空间下的内容 
        //使用NPOI.HSSF.UserModel空间下的HSSFWorkbook操作Excel2003        
        //使用NPOI.XSSF.UserModel空间下的XSSFWorkbook操作Excel2007
        //"application/vnd.openxmlformats-officedocument.wordprocessingml.document" (for .docx files)
        //"application/vnd.openxmlformats-officedocument.wordprocessingml.template" (for .dotx files)
        //"application/vnd.openxmlformats-officedocument.presentationml.presentation" (for .pptx files)
        //"application/vnd.openxmlformats-officedocument.presentationml.slideshow" (for .ppsx files)
        //"application/vnd.openxmlformats-officedocument.presentationml.template" (for .potx files)
        //"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" (for .xlsx files)
        //"application/vnd.openxmlformats-officedocument.spreadsheetml.template" (for .xltx files)
        // 相对于Office2003是这样的
        // Response.ContentType = "application/vnd.ms-excel

        protected void Page_Load(object sender, EventArgs e)
        {

        }

        #region  公共方法

        //返回一个初始化的Excel 2003对象 
        private HSSFWorkbook InitializeWorkbook_HSSF()
        {
            HSSFWorkbook workBook = new HSSFWorkbook();

            //创建版本信息 右键--属性--详细信息--来源--公司
            DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
            dsi.Company = "Test Company";
            workBook.DocumentSummaryInformation = dsi;

            //右键--属性--详细信息--说明--主题
            SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
            si.Subject = "Test HyperLink In  Excel";
            workBook.SummaryInformation = si;

            return workBook;
        }

        //下载一个Excel 2003文件        
        private void downLoadFile_HSSF(HSSFWorkbook workBook, string fileName)
        {
            //设置响应的类型为Excel
            Response.ContentType = "application/vnd.ms-excel";
            //设置下载的Excel文件名
            Response.AddHeader("Content-Disposition", string.Format("attachment;filename={0}", fileName));//目前只能为03的Excel
            //Clear方法删除所有缓存中的HTML输出。但此方法只删除Response显示输入信息，不删除Response头信息。以免影响导出数据的完整性。
            Response.Clear();

            //注意类型
            using (MemoryStream file = new MemoryStream())
            {
                //将工作簿的内容放到内存流中
                workBook.Write(file);
                //将内存流转换成字节数组发送到客户端
                /*
                 * //这种写法和下面的相比 会产生大量的个byte[]临时变量数据
                Response.BinaryWrite(file.GetBuffer());
                Response.End();
                */
                file.WriteTo(Response.OutputStream);
            }
        }
        
        private void SetCellWidth(ISheet sheet,int cellIndex, ICell cell,int cellWidth = 0)
        {
            if(cellWidth == 0)
            {
                cellWidth = Encoding.Default.GetBytes(cell.ToString()).Length * 300;
            }          
            sheet1.SetColumnWidth(cellIndex, cellWidth);
        }

        //创建style
        private Dictionary<String, ICellStyle> createStyles(IWorkbook wb)
        {
            Dictionary<String, ICellStyle> styles = new Dictionary<String, ICellStyle>();
            IDataFormat df = wb.CreateDataFormat();

            ICellStyle style;
            IFont headerFont = wb.CreateFont();
            headerFont.Boldweight = (short)(FontBoldWeight.Bold);
            style = CreateBorderedStyle(wb);
            style.Alignment = HorizontalAlignment.Center;
            style.FillForegroundColor = (IndexedColors.LightCornflowerBlue.Index);
            style.FillPattern = FillPattern.SolidForeground;
            style.SetFont(headerFont);
            styles.Add("header", style);

            style = CreateBorderedStyle(wb);
            style.Alignment = HorizontalAlignment.Center;
            style.FillForegroundColor = (IndexedColors.LightCornflowerBlue.Index);
            style.FillPattern = FillPattern.SolidForeground;
            style.SetFont(headerFont);
            style.DataFormat = (df.GetFormat("d-mmm"));
            styles.Add("header_date", style);

            IFont font1 = wb.CreateFont();
            font1.Boldweight = (short)(FontBoldWeight.Bold);
            style = CreateBorderedStyle(wb);
            style.Alignment = HorizontalAlignment.Center;
            style.SetFont(font1);
            styles.Add("cell_b", style);

            style = CreateBorderedStyle(wb);
            style.Alignment = HorizontalAlignment.Center;
            style.SetFont(font1);
            styles.Add("cell_b_centered", style);

            style = CreateBorderedStyle(wb);
            style.Alignment = HorizontalAlignment.Center;
            style.SetFont(font1);
            style.DataFormat = (df.GetFormat("d-mmm"));
            styles.Add("cell_b_date", style);

            style = CreateBorderedStyle(wb);
            style.Alignment = HorizontalAlignment.Center;
            style.SetFont(font1);
            style.FillForegroundColor = (IndexedColors.Grey25Percent.Index);
            style.FillPattern = FillPattern.SolidForeground;
            style.DataFormat = (df.GetFormat("d-mmm"));
            styles.Add("cell_g", style);

            IFont font2 = wb.CreateFont();
            font2.Color = (IndexedColors.Blue.Index);
            font2.Boldweight = (short)(FontBoldWeight.Bold);
            style = CreateBorderedStyle(wb);
            style.Alignment = HorizontalAlignment.Center;
            style.SetFont(font2);
            styles.Add("cell_bb", style);

            style = CreateBorderedStyle(wb);
            style.Alignment = HorizontalAlignment.Center;
            style.SetFont(font1);
            style.FillForegroundColor = (IndexedColors.Grey25Percent.Index);
            style.FillPattern = FillPattern.SolidForeground;
            style.DataFormat = (df.GetFormat("d-mmm"));
            styles.Add("cell_bg", style);

            IFont font3 = wb.CreateFont();
            font3.FontHeightInPoints = ((short)14);
            font3.Color = (IndexedColors.DarkBlue.Index);
            font3.Boldweight = (short)(FontBoldWeight.Bold);
            style = CreateBorderedStyle(wb);
            style.Alignment = HorizontalAlignment.Center;
            style.SetFont(font3);
            style.WrapText = (true);
            styles.Add("cell_h", style);

            style = CreateBorderedStyle(wb);
            style.Alignment = HorizontalAlignment.Center;
            style.WrapText = (true);
            styles.Add("cell_normal", style);

            style = CreateBorderedStyle(wb);
            style.Alignment = HorizontalAlignment.Center;
            style.WrapText = (true);
            styles.Add("cell_normal_centered", style);

            style = CreateBorderedStyle(wb);
            style.Alignment = HorizontalAlignment.Center;
            style.WrapText = (true);
            style.DataFormat = (df.GetFormat("d-mmm"));
            styles.Add("cell_normal_date", style);

            style = CreateBorderedStyle(wb);
            style.Alignment = HorizontalAlignment.Center;
            style.Indention = ((short)1);
            style.WrapText = (true);
            styles.Add("cell_indented", style);

            style = CreateBorderedStyle(wb);
            style.FillForegroundColor = (IndexedColors.Blue.Index);
            style.FillPattern = FillPattern.SolidForeground;
            styles.Add("cell_blue", style);

            return styles;
        }

        private Dictionary<String, ICellStyle> createStyles_Calendar(IWorkbook wb)
        {
            Dictionary<String, ICellStyle> styles = new Dictionary<String, ICellStyle>();

            short borderColor = IndexedColors.Grey50Percent.Index;

            ICellStyle style;
            IFont titleFont = wb.CreateFont();
            titleFont.FontHeightInPoints = ((short)48);
            titleFont.Color = (IndexedColors.DarkBlue.Index);
            style = wb.CreateCellStyle();
            style.Alignment = (HorizontalAlignment.Center);
            style.VerticalAlignment = VerticalAlignment.Center;
            style.SetFont(titleFont);
            styles.Add("title", style);

            IFont monthFont = wb.CreateFont();
            monthFont.FontHeightInPoints = ((short)12);
            monthFont.Color = (IndexedColors.White.Index);
            monthFont.Boldweight = (short)(FontBoldWeight.Bold);
            style = wb.CreateCellStyle();
            style.Alignment = (HorizontalAlignment.Center);
            style.VerticalAlignment = (VerticalAlignment.Center);
            style.FillForegroundColor = (IndexedColors.DarkBlue.Index);
            style.FillPattern = (FillPattern.SolidForeground);
            style.SetFont(monthFont);
            styles.Add("month", style);

            IFont dayFont = wb.CreateFont();
            dayFont.FontHeightInPoints = ((short)14);
            dayFont.Boldweight = (short)(FontBoldWeight.Bold);
            style = wb.CreateCellStyle();
            style.Alignment = (HorizontalAlignment.Left);
            style.VerticalAlignment = (VerticalAlignment.Top);
            style.FillForegroundColor = (IndexedColors.LightCornflowerBlue.Index);
            style.FillPattern = (FillPattern.SolidForeground);
            style.BorderLeft = (NPOI.SS.UserModel.BorderStyle.Thin);
            style.LeftBorderColor = (borderColor);
            style.BorderBottom = (NPOI.SS.UserModel.BorderStyle.Thin);
            style.BottomBorderColor = (borderColor);
            style.SetFont(dayFont);
            styles.Add("weekend_left", style);

            style = wb.CreateCellStyle();
            style.Alignment = (HorizontalAlignment.Center);
            style.VerticalAlignment = (VerticalAlignment.Top);
            style.FillForegroundColor = (IndexedColors.LightCornflowerBlue.Index);
            style.FillPattern = (FillPattern.SolidForeground);
            style.BorderRight = (NPOI.SS.UserModel.BorderStyle.Thin);
            style.RightBorderColor = (borderColor);
            style.BorderBottom = (NPOI.SS.UserModel.BorderStyle.Thin);
            style.BottomBorderColor = (borderColor);
            styles.Add("weekend_right", style);

            style = wb.CreateCellStyle();
            style.Alignment = (HorizontalAlignment.Left);
            style.VerticalAlignment = (VerticalAlignment.Top);
            style.BorderLeft = (NPOI.SS.UserModel.BorderStyle.Thin);
            style.FillForegroundColor = (IndexedColors.White.Index);
            style.FillPattern = (FillPattern.SolidForeground);
            style.LeftBorderColor = (borderColor);
            style.BorderBottom = (NPOI.SS.UserModel.BorderStyle.Thin);
            style.BottomBorderColor = (borderColor);
            style.SetFont(dayFont);
            styles.Add("workday_left", style);

            style = wb.CreateCellStyle();
            style.Alignment = (HorizontalAlignment.Center);
            style.VerticalAlignment = (VerticalAlignment.Top);
            style.FillForegroundColor = (IndexedColors.White.Index);
            style.FillPattern = (FillPattern.SolidForeground);
            style.BorderRight = (NPOI.SS.UserModel.BorderStyle.Thin);
            style.RightBorderColor = (borderColor);
            style.BorderBottom = (NPOI.SS.UserModel.BorderStyle.Thin);
            style.BottomBorderColor = (borderColor);
            styles.Add("workday_right", style);

            style = wb.CreateCellStyle();
            style.BorderLeft = (NPOI.SS.UserModel.BorderStyle.Thin);
            style.FillForegroundColor = (IndexedColors.Grey25Percent.Index);
            style.FillPattern = (FillPattern.SolidForeground);
            style.BorderBottom = (NPOI.SS.UserModel.BorderStyle.Thin);
            style.BottomBorderColor = (borderColor);
            styles.Add("grey_left", style);

            style = wb.CreateCellStyle();
            style.FillForegroundColor = (IndexedColors.Grey25Percent.Index);
            style.FillPattern = (FillPattern.SolidForeground);
            style.BorderRight = (NPOI.SS.UserModel.BorderStyle.Thin);
            style.RightBorderColor = (borderColor);
            style.BorderBottom = (NPOI.SS.UserModel.BorderStyle.Thin);
            style.BottomBorderColor = (borderColor);
            styles.Add("grey_right", style);

            return styles;
        }

        //创建边框style
        private ICellStyle CreateBorderedStyle(IWorkbook wb)
        {
            ICellStyle style = wb.CreateCellStyle();
            style.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            style.RightBorderColor = (IndexedColors.Black.Index);
            style.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            style.BottomBorderColor = (IndexedColors.Black.Index);
            style.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            style.LeftBorderColor = (IndexedColors.Black.Index);
            style.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            style.TopBorderColor = (IndexedColors.Black.Index);
            return style;
        }

        //Excel 2007导出
        private void downLoadFile_XSSF(IWorkbook workBook, string fileName)
        {
            //设置响应的类型为Excel
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            //设置下载的Excel文件名
            Response.AddHeader("Content-Disposition", string.Format("attachment;filename={0}", fileName));
            //Clear方法删除所有缓存中的HTML输出。但此方法只删除Response显示输入信息，不删除Response头信息。以免影响导出数据的完整性。
            Response.Clear();



            //注意类型
            //using (MemoryStream file = new MemoryStream())
            //{
            //    //将工作簿的内容放到内存流中
            //    workBook.Write(file);                
            //    //file.WriteTo(Response.OutputStream);
            //    Response.Flush();
            //    Response.End();
            //}

            //目前来说只有这个方法 上面的方法不行
            using (FileStream f = File.Create(@"c:\test.xlsx"))
            {
                workBook.Write(f);
            }
            Response.WriteFile(@"c:\test.xlsx");
            Response.Flush();
            System.IO.File.Delete(@"c:\test.xlsx");//删除保存的文件
            Response.End();
        }

        //Excel 2007导出
        private void downLoadFile_XWPF(XWPFDocument doc, string fileName)
        {
            //设置响应的类型为word
            Response.ContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
            //设置下载的Excel文件名
            Response.AddHeader("Content-Disposition", string.Format("attachment;filename={0}", fileName));
            //Clear方法删除所有缓存中的HTML输出。但此方法只删除Response显示输入信息，不删除Response头信息。以免影响导出数据的完整性。
            Response.Clear();



            //注意类型
            //using (MemoryStream file = new MemoryStream())
            //{
            //    //将工作簿的内容放到内存流中
            //    workBook.Write(file);                
            //    //file.WriteTo(Response.OutputStream);
            //    Response.Flush();
            //    Response.End();
            //}

            //目前来说只有这个方法 上面的方法不行
            using (FileStream f = File.Create(@"c:\test.docx"))
            {
                doc.Write(f);
            }
            Response.WriteFile(@"c:\test.docx");
            Response.Flush();
            System.IO.File.Delete(@"c:\test.docx");//删除保存的文件
            Response.End();
        }

        #endregion

        #region Excel 2003

        //创建一个空白的Excel
        protected void lnkbtnExcel_2003_EmptyExcel_Click(object sender, EventArgs e)
        {
            HSSFWorkbook workBook = InitializeWorkbook_HSSF();
            workBook.CreateSheet("Sheet1");
            workBook.CreateSheet("Sheet2");
            workBook.CreateSheet("Sheet3");
            workBook.CreateSheet("Sheet4");
            ((HSSFSheet)workBook.GetSheetAt(0)).AlternativeFormula = false;
            ((HSSFSheet)workBook.GetSheetAt(0)).AlternativeExpression = false;
            downLoadFile_HSSF(workBook, "Excel_2003_EmptyExcel.xls");
        }

        //带有HyperLink的Excel
        protected void lnkbtnExcel_2003_HyperLink_Click(object sender, EventArgs e)
        {
            HSSFWorkbook workBook = new HSSFWorkbook();
            workBook = InitializeWorkbook_HSSF();

            ICellStyle link_style = workBook.CreateCellStyle();
            IFont hlink_font = workBook.CreateFont();
            hlink_font.Underline = FontUnderlineType.Single;
            hlink_font.Color = HSSFColor.Blue.Index;
            link_style.SetFont(hlink_font);

            ICell cell;
            ISheet sheet = workBook.CreateSheet("Hyperlinks");

            //URL
            cell = sheet.CreateRow(0).CreateCell(0);
            cell.SetCellValue("URL Link");
            HSSFHyperlink link = new HSSFHyperlink(HyperlinkType.Url);
            link.Address = ("http://poi.apache.org/");
            cell.Hyperlink = (link);
            cell.CellStyle = (link_style);

            //link to a file in the current directory
            cell = sheet.CreateRow(1).CreateCell(0);
            cell.SetCellValue("File Link");
            link = new HSSFHyperlink(HyperlinkType.File);
            link.Address = ("link1.xls");
            cell.Hyperlink = (link);
            cell.CellStyle = (link_style);

            //e-mail link
            cell = sheet.CreateRow(2).CreateCell(0);
            cell.SetCellValue("Email Link");
            link = new HSSFHyperlink(HyperlinkType.Email);
            //note, if subject contains white spaces, make sure they are url-encoded
            link.Address = ("mailto:poi@apache.org?subject=Hyperlinks");
            cell.Hyperlink = (link);
            cell.CellStyle = (link_style);

            //link to a place in this workbook

            //Create a target sheet and cell
            ISheet sheet2 = workBook.CreateSheet("Target ISheet");
            sheet2.CreateRow(0).CreateCell(0).SetCellValue("Target ICell");

            cell = sheet.CreateRow(3).CreateCell(0);
            cell.SetCellValue("Worksheet Link");
            link = new HSSFHyperlink(HyperlinkType.Document);
            link.Address = ("'Target ISheet'!A1");
            cell.Hyperlink = (link);
            cell.CellStyle = (link_style);

            downLoadFile_HSSF(workBook, "Excel_2003_HyperLink.xls");
        }

        //带有字体应用的Excel
        protected void lnkbtnExcel_2003_Font_Click(object sender, EventArgs e)
        {
            HSSFWorkbook workBook = InitializeWorkbook_HSSF();
            ISheet sheet1 = workBook.CreateSheet("Sheet1");

            //font style1: underlined, italic, red color, fontsize=20
            IFont font1 = workBook.CreateFont();
            font1.Color = HSSFColor.Red.Index;
            font1.IsItalic = true;
            font1.Underline = FontUnderlineType.Double;
            font1.FontHeightInPoints = 20;

            //bind font with style 1
            ICellStyle style1 = workBook.CreateCellStyle();
            style1.SetFont(font1);

            //font style2: strikeout line, green color, fontsize=15, fontname='宋体'
            IFont font2 = workBook.CreateFont();
            font2.Color = HSSFColor.OliveGreen.Index;
            font2.IsStrikeout = true;
            font2.FontHeightInPoints = 15;
            font2.FontName = "宋体";

            //bind font with style 2
            ICellStyle style2 = workBook.CreateCellStyle();
            style2.SetFont(font2);

            //apply font styles
            ICell cell1 = HSSFCellUtil.CreateCell(sheet1.CreateRow(1), 1, "Hello World!");
            cell1.CellStyle = style1;
            ICell cell2 = HSSFCellUtil.CreateCell(sheet1.CreateRow(3), 1, "早上好！");
            cell2.CellStyle = style2;

            //cell with rich text 
            ICell cell3 = sheet1.CreateRow(5).CreateCell(1);
            HSSFRichTextString richtext = new HSSFRichTextString("Microsoft OfficeTM");

            //apply font to "Microsoft Office"
            IFont font4 = workBook.CreateFont();
            font4.FontHeightInPoints = 12;
            richtext.ApplyFont(0, 16, font4);
            //apply font to "TM"
            IFont font3 = workBook.CreateFont();
            font3.TypeOffset = FontSuperScript.Super;
            font3.IsItalic = true;
            font3.Color = HSSFColor.Blue.Index;
            font3.FontHeightInPoints = 8;
            richtext.ApplyFont(16, 18, font3);

            cell3.SetCellValue(richtext);

            downLoadFile_HSSF(workBook, "Excel_2003_Font.xls");
        }

        //带有自动适应列宽的Excel
        protected void lnkbtnExcel_2003_AutoWidth_Click(object sender, EventArgs e)
        {
            //创建DataTable 将数据库中没有的数据放到这个DT中
            DataTable dt = new DataTable();
            dt.Columns.Add("列1", typeof(string));
            dt.Columns.Add("列2", typeof(string));
            dt.Columns.Add("列3", typeof(string));
            //创建DatatTable 结束---------------------------

            //开始给临时datatable赋值
            for (int i = 0; i < 10; i++)
            {
                DataRow row = dt.NewRow();
                row["列1"] = "列111111111111111111111111111111";
                row["列2"] = "列222222222222222222222222222222222222222";
                row["列3"] = "列3333333322222222222211111111111111111111111113";
                dt.Rows.Add(row);
            }


            HSSFWorkbook workBook = InitializeWorkbook_HSSF();
            ISheet sheet = workBook.CreateSheet("sheet1");
            //头部标题
            IRow headerRow = sheet.CreateRow(0);
            //循环添加标题
            foreach (DataColumn column in dt.Columns)
                headerRow.CreateCell(column.Ordinal).SetCellValue(column.ColumnName);

            // 内容
            int paymentRowIndex = 1;

            foreach (DataRow row in dt.Rows)
            {
                IRow newRow = sheet.CreateRow(paymentRowIndex);

                //循环添加列的对应内容
                foreach (DataColumn column in dt.Columns)
                {
                    newRow.CreateCell(column.Ordinal).SetCellValue(row[column].ToString());
                }

                paymentRowIndex++;
            }

            //列宽自适应，只对英文和数字有效
            for (int i = 0; i <= dt.Rows.Count; i++)
            {
                sheet.AutoSizeColumn(i);
            }

            //获取当前列的宽度，然后对比本列的长度，取最大值
            for (int columnNum = 0; columnNum <= dt.Columns.Count; columnNum++)
            {
                int columnWidth = sheet.GetColumnWidth(columnNum) / 256;
                for (int rowNum = 1; rowNum <= sheet.LastRowNum; rowNum++)
                {
                    IRow currentRow;
                    //当前行未被使用过
                    if (sheet.GetRow(rowNum) == null)
                    {
                        currentRow = sheet.CreateRow(rowNum);
                    }
                    else
                    {
                        currentRow = sheet.GetRow(rowNum);
                    }

                    if (currentRow.GetCell(columnNum) != null)
                    {
                        ICell currentCell = currentRow.GetCell(columnNum);
                        int length = Encoding.Default.GetBytes(currentCell.ToString()).Length;
                        if (columnWidth < length)
                        {
                            columnWidth = length;
                        }
                    }
                }
                sheet.SetColumnWidth(columnNum, columnWidth * 256);
            }

            downLoadFile_HSSF(workBook, "Excel_2003_AutoWidth");
        }

        //包含业务计划的Excel 支持2007
        protected void lnkbtnExcel_2003_BusinessPlan_Click(object sender, EventArgs e)
        {
            SimpleDateFormat fmt = new SimpleDateFormat("dd-MMM");
            String[] titles = { "ID", "Project Name", "Owner", "Days", "Start", "End" };
            String[][] data = new string[18][];

            data[0] = new string[] {"1.0", "Marketing Research Tactical Plan", "J. Dow", "70", "9-Jul", null,
                "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x"};
            data[1] = new string[] { null };
            data[2] = new string[] {"1.1", "Scope Definition Phase", "J. Dow", "10", "9-Jul", null,
                "x", "x", null, null,  null, null, null, null, null, null, null};
            data[3] = new string[] {"1.1.1", "Define research objectives", "J. Dow", "3", "9-Jul", null,
                    "x", null, null, null,  null, null, null, null, null, null, null};
            data[4] = new string[] {"1.1.2", "Define research requirements", "S. Jones", "7", "10-Jul", null,
                "x", "x", null, null,  null, null, null, null, null, null, null};
            data[5] = new string[] {"1.1.3", "Determine in-house resource or hire vendor", "J. Dow", "2", "15-Jul", null,
                "x", "x", null, null,  null, null, null, null, null, null, null};
            data[6] = new string[] { null };
            data[7] = new string[] {"1.2", "Vendor Selection Phase", "J. Dow", "19", "19-Jul", null,
                null, "x", "x", "x",  "x", null, null, null, null, null, null};
            data[8] = new string[] {"1.2.1", "Define vendor selection criteria", "J. Dow", "3", "19-Jul", null,
                null, "x", null, null,  null, null, null, null, null, null, null};
            data[9] = new string[] {"1.2.2", "Develop vendor selection questionnaire", "S. Jones, T. Wates", "2", "22-Jul", null,
                null, "x", "x", null,  null, null, null, null, null, null, null};
            data[10] = new string[] {"1.2.3", "Develop Statement of Work", "S. Jones", "4", "26-Jul", null,
                null, null, "x", "x",  null, null, null, null, null, null, null};
            data[11] = new string[] {"1.2.4", "Evaluate proposal", "J. Dow, S. Jones", "4", "2-Aug", null,
                null, null, null, "x",  "x", null, null, null, null, null, null};
            data[12] = new string[] {"1.2.5", "Select vendor", "J. Dow", "1", "6-Aug", null,
                null, null, null, null,  "x", null, null, null, null, null, null};
            data[13] = new string[] { null };
            data[14] = new string[] {"1.3", "Research Phase", "G. Lee", "47", "9-Aug", null,
                null, null, null, null,  "x", "x", "x", "x", "x", "x", "x"};
            data[15] = new string[] {"1.3.1", "Develop market research information needs questionnaire", "G. Lee", "2", "9-Aug", null,
                null, null, null, null,  "x", null, null, null, null, null, null};
            data[16] = new string[] {"1.3.2", "Interview marketing group for market research needs", "G. Lee", "2", "11-Aug", null,
                null, null, null, null,  "x", "x", null, null, null, null, null};
            data[17] = new string[] {"1.3.3", "Document information needs", "G. Lee, S. Jones", "1", "13-Aug", null,
                null, null, null, null,  null, "x", null, null, null, null, null};

            HSSFWorkbook workBook = InitializeWorkbook_HSSF();
            Dictionary<String, ICellStyle> styles = createStyles(workBook);

            ISheet sheet = workBook.CreateSheet("Business Plan");

            //turn off gridlines
            sheet.DisplayGridlines = (false);
            sheet.IsPrintGridlines = (false);
            sheet.FitToPage = (true);
            sheet.HorizontallyCenter = (true);
            IPrintSetup printSetup = sheet.PrintSetup;
            printSetup.Landscape = (true);

            //the following three statements are required only for HSSF
            sheet.Autobreaks = (true);
            printSetup.FitHeight = ((short)1);
            printSetup.FitWidth = ((short)1);

            //the header row: centered text in 48pt font
            IRow headerRow = sheet.CreateRow(0);
            headerRow.HeightInPoints = (12.75f);
            for (int i = 0; i < titles.Length; i++)
            {
                ICell cell = headerRow.CreateCell(i);
                cell.SetCellValue(titles[i]);
                cell.CellStyle = (styles["header"]);
            }
            //columns for 11 weeks starting from 9-Jul
            DateTime dt = new DateTime(DateTime.Now.Year, 6, 9);
            for (int i = 0; i < 11; i++)
            {
                ICell cell = headerRow.CreateCell(titles.Length + i);
                cell.SetCellValue(dt);
                cell.CellStyle = (styles[("header_date")]);
                //calendar.roll(Calendar.WEEK_OF_YEAR, true);
                dt.AddDays(7);
            }
            //freeze the first row
            sheet.CreateFreezePane(0, 1);

            IRow row;
            //ICell cell;
            int rownum = 1;
            for (int i = 0; i < data.Length; i++, rownum++)
            {
                row = sheet.CreateRow(rownum);
                if (data[i] == null) continue;

                for (int j = 0; j < data[i].Length; j++)
                {
                    ICell cell = row.CreateCell(j);
                    String styleName;
                    bool isHeader = i == 0 || data[i - 1] == null;
                    switch (j)
                    {
                        case 0:
                            if (isHeader)
                            {
                                styleName = "cell_b";
                                cell.SetCellValue(Double.Parse(data[i][j]));
                            }
                            else
                            {
                                styleName = "cell_normal";
                                cell.SetCellValue(data[i][j]);
                            }
                            break;
                        case 1:
                            if (isHeader)
                            {
                                styleName = i == 0 ? "cell_h" : "cell_bb";
                            }
                            else
                            {
                                styleName = "cell_indented";
                            }
                            cell.SetCellValue(data[i][j]);
                            break;
                        case 2:
                            styleName = isHeader ? "cell_b" : "cell_normal";
                            cell.SetCellValue(data[i][j]);
                            break;
                        case 3:
                            styleName = isHeader ? "cell_b_centered" : "cell_normal_centered";
                            cell.SetCellValue(int.Parse(data[i][j]));
                            break;
                        case 4:
                            {
                                //calendar.setTime(fmt.parse(data[i][j]));
                                //calendar.set(Calendar.YEAR, year);

                                DateTime dt2 = DateTime.Parse(DateTime.Now.Year.ToString() + "-" + data[i][j]);

                                cell.SetCellValue(dt2);
                                styleName = isHeader ? "cell_b_date" : "cell_normal_date";
                                break;
                            }
                        case 5:
                            {
                                int r = rownum + 1;
                                String fmla = "IF(AND(D" + r + ",E" + r + "),E" + r + "+D" + r + ",\"\")";
                                cell.SetCellFormula(fmla);
                                styleName = isHeader ? "cell_bg" : "cell_g";
                                break;
                            }
                        default:
                            styleName = data[i][j] != null ? "cell_blue" : "cell_normal";
                            break;
                    }

                    cell.CellStyle = (styles[(styleName)]);
                }
            }

            //group rows for each phase, row numbers are 0-based
            sheet.GroupRow(4, 6);
            sheet.GroupRow(9, 13);
            sheet.GroupRow(16, 18);

            //set column widths, the width is measured in units of 1/256th of a character width
            sheet.SetColumnWidth(0, 256 * 6);
            sheet.SetColumnWidth(1, 256 * 33);
            sheet.SetColumnWidth(2, 256 * 20);
            sheet.SetZoom(3, 4);

            downLoadFile_HSSF(workBook, "Excel_2003_BusinessPlan");
        }

        //包含日历的Excel 支持2007
        protected void lnkbtnExcel_2003_Calendar_Click(object sender, EventArgs e)
        {
            String[] days = { "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday" };
            String[] months = { "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December" };

            DateTime dt = DateTime.Now;
            HSSFWorkbook workBook = new HSSFWorkbook();

            Dictionary<String, ICellStyle> styles = createStyles_Calendar(workBook);
            DateTime dtM;
            for (int month = 0; month < 12; month++)
            {
                dtM = new DateTime(dt.Year, month + 1, 1);
                //calendar.set(Calendar.MONTH, month);
                //calendar.set(Calendar.DAY_OF_MONTH, 1);
                //create a sheet for each month
                ISheet sheet = workBook.CreateSheet(months[month]);

                //turn off gridlines
                sheet.DisplayGridlines = (false);
                sheet.IsPrintGridlines = (false);
                sheet.FitToPage = (true);
                sheet.HorizontallyCenter = (true);
                IPrintSetup printSetup = sheet.PrintSetup;
                printSetup.Landscape = (true);//横向打印

                //the following three statements are required only for HSSF
                sheet.Autobreaks = (true);
                printSetup.FitHeight = ((short)1);
                printSetup.FitWidth = ((short)1);

                //the header row: centered text in 48pt font
                IRow headerRow = sheet.CreateRow(0);
                headerRow.HeightInPoints = (80);
                ICell titleCell = headerRow.CreateCell(0);
                titleCell.SetCellValue(months[month] + " " + dt.Year);
                titleCell.CellStyle = (styles[("title")]);
                sheet.AddMergedRegion(CellRangeAddress.ValueOf("$A$1:$N$1"));

                //header with month titles
                IRow monthRow = sheet.CreateRow(1);
                for (int i = 0; i < days.Length; i++)
                {
                    //set column widths, the width is measured in units of 1/256th of a character width

                    //奇数列的宽度
                    sheet.SetColumnWidth(i * 2, 5 * 256); //the column is 5 characters wide
                    //偶数列的宽度
                    sheet.SetColumnWidth(i * 2 + 1, 13 * 256); //the column is 13 characters wide
                    //合并单元格 第二行中的i * 2位置到第二行的i * 2+1的位置
                    sheet.AddMergedRegion(new CellRangeAddress(1, 1, i * 2, i * 2 + 1));
                    //把值赋值给第一个列（两个列合并的单元格）
                    ICell monthCell = monthRow.CreateCell(i * 2);
                    monthCell.SetCellValue(days[i]);
                    monthCell.CellStyle = (styles["month"]);
                }

                int cnt = 0;//表示周几  0表示周天 1-6 对应周一-周六
                int day = 1;
                int rownum = 2;//所在行
                for (int j = 0; j < 6; j++)
                {
                    IRow row = sheet.CreateRow(rownum);
                    row.HeightInPoints = (100);
                    for (int i = 0; i < days.Length; i++)
                    {
                        ICell dayCell_1 = row.CreateCell(i * 2);
                        ICell dayCell_2 = row.CreateCell(i * 2 + 1);
                        int day_of_week = (int)dtM.DayOfWeek;//这个主要是定位1号的时候 是具体的星期几 前面有几天的空白                        
                        if (cnt >= day_of_week && dtM.Month == (month + 1))
                        {
                            dayCell_1.SetCellValue(day);
                            //calendar.set(Calendar.DAY_OF_MONTH, ++day);
                            if (day == 29)
                            {

                            }
                            day++;
                            dtM = dtM.AddDays(1);

                            if (i == 0 || i == days.Length - 1)
                            {
                                dayCell_1.CellStyle = (styles["weekend_left"]);
                                dayCell_2.CellStyle = (styles["weekend_right"]);
                            }
                            else
                            {
                                dayCell_1.CellStyle = (styles["workday_left"]);
                                dayCell_2.CellStyle = (styles["workday_right"]);
                            }
                        }
                        else
                        {
                            dayCell_1.CellStyle = (styles["grey_left"]);
                            dayCell_2.CellStyle = (styles["grey_right"]);
                        }
                        cnt++;
                    }
                    rownum++;
                    if (dtM.Month > (month + 1)) break;
                }
            }
            downLoadFile_HSSF(workBook, "Excel_2003_Calendar");
        }

        //可以改变sheet颜色的Excel
        protected void lnkbtnExcel_2003_SheetColor_Click(object sender, EventArgs e)
        {
            HSSFWorkbook workBook = InitializeWorkbook_HSSF();
            ISheet sheet1 = workBook.CreateSheet("Sheet1");
            sheet1.TabColorIndex = HSSFColor.Red.Index;
            ISheet sheet2 = workBook.CreateSheet("Sheet2");
            sheet2.TabColorIndex = HSSFColor.Blue.Index;
            ISheet sheet3 = workBook.CreateSheet("Sheet3");
            sheet3.TabColorIndex = HSSFColor.Aqua.Index;
            downLoadFile_HSSF(workBook, "Excel_2003_SheetColor.xls");
        }

        //带有彩色矩阵的Excel
        protected void lnkbtnExcel_2003_ColorMatrixTable_Click(object sender, EventArgs e)
        {
            HSSFWorkbook workBook = InitializeWorkbook_HSSF();
            ISheet sheet1 = workBook.CreateSheet("Sheet1");

            int x = 1;
            for (int i = 0; i < 15; i++)
            {
                IRow row = sheet1.CreateRow(i);
                for (int j = 0; j < 15; j++)
                {
                    ICell cell = row.CreateCell(j);
                    if (x % 2 == 0)
                    {
                        //fill background with blue
                        ICellStyle style1 = workBook.CreateCellStyle();
                        style1.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Blue.Index2;
                        style1.FillPattern = FillPattern.SolidForeground;
                        cell.CellStyle = style1;
                    }
                    else
                    {
                        //fill background with yellow
                        ICellStyle style1 = workBook.CreateCellStyle();
                        style1.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Yellow.Index2;
                        style1.FillPattern = FillPattern.SolidForeground;
                        cell.CellStyle = style1;
                    }
                    x++;
                }
            }

            downLoadFile_HSSF(workBook, "Excel_2003_ColorMatrixTable.xls");
        }

        //带有规则格式的Excel
        protected void lnkbtnExcel_2003_ConditionalFormat_Click(object sender, EventArgs e)
        {
            HSSFWorkbook workBook = InitializeWorkbook_HSSF();

            ISheet sheet1 = workBook.CreateSheet("Sheet1");
            ISheetConditionalFormatting hscf = sheet1.SheetConditionalFormatting;

            // Define a Conditional Formatting rule, which triggers formatting
            // when cell's value is bigger than 55 and smaller than 500
            // applies patternFormatting defined below.
            IConditionalFormattingRule rule = hscf.CreateConditionalFormattingRule(
                ComparisonOperator.Between,
                "55", // 1st formula 
                "500"     // 2nd formula 
            );

            // Create pattern with red background
            IPatternFormatting patternFmt = rule.CreatePatternFormatting();
            patternFmt.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.Red.Index;

            //// Define a region containing first column
            CellRangeAddress[] regions = {
                new CellRangeAddress(0, 65535,0,1) //表示有效范围在0-65535行的A和B列有效
            };
            // Apply Conditional Formatting rule defined above to the regions  
            hscf.AddConditionalFormatting(regions, rule);

            //fill cell with numeric values
            sheet1.CreateRow(0).CreateCell(0).SetCellValue(50);
            sheet1.CreateRow(0).CreateCell(1).SetCellValue(101);
            sheet1.CreateRow(1).CreateCell(1).SetCellValue(25);
            sheet1.CreateRow(1).CreateCell(0).SetCellValue(150);

            downLoadFile_HSSF(workBook, "Excel_2003_ConditionalFormat.xls");
        }

        //把Excel转换成Html
        protected void lnkbtnExcel_2003_ExcelToHtml_Click(object sender, EventArgs e)
        {
            HSSFWorkbook workBook = InitializeWorkbook_HSSF();
            //获取文件
            string fileName = Server.MapPath("ExcelMode/ExcelToHtml.xls");

            workBook = ExcelToHtmlUtils.LoadXls(fileName);
            ExcelToHtmlConverter excelToHtmlConverter = new ExcelToHtmlConverter();

            //set output parameter
            excelToHtmlConverter.OutputColumnHeaders = false;
            excelToHtmlConverter.OutputHiddenColumns = true;
            excelToHtmlConverter.OutputHiddenRows = true;
            excelToHtmlConverter.OutputLeadingSpacesAsNonBreaking = false;
            excelToHtmlConverter.OutputRowNumbers = true;
            excelToHtmlConverter.UseDivsToSpan = true;

            //process the excel file
            excelToHtmlConverter.ProcessWorkbook(workBook);

            //output the html file  保存到Server.MapPath("ExcelMode/ExcelToHtml.xls");
            excelToHtmlConverter.Document.Save(Path.ChangeExtension(fileName, "html"));
        }

        //复制别的Excel的行和列产生新的Excel
        protected void lnkbtnExcel_2003_CopyRowsAndCell_Click(object sender, EventArgs e)
        {
            string path = Server.MapPath("ExcelMode/CopyRowsAndCell.xls");
            FileStream file = File.OpenRead(path);
            HSSFWorkbook workBook = new HSSFWorkbook(file);
            ISheet s = workBook.GetSheetAt(0);

            ICell cell = s.GetRow(4).GetCell(1);
            cell.CopyCellTo(3); //copy B5 to D5

            IRow c = s.GetRow(3);
            c.CopyCell(0, 1);   //copy A4 to B4

            s.CopyRow(0, 1);     //copy row A to row B, original row B will be moved to row C automatically

            downLoadFile_HSSF(workBook, "Excel_2003_CopyRowsAndCell.xls");
        }

        //复制一个Excel的Sheet产生新的Excel(用于把多个Excel合成一个Excel)
        protected void lnkbtnExcel_2003_CopySheet_Click(object sender, EventArgs e)
        {
            string path = Server.MapPath("ExcelMode/CopySheet1.xls");
            FileStream file = File.OpenRead(path);
            HSSFWorkbook workBook1 = new HSSFWorkbook(file);//2003 所以选择文件的时候 要选择03的Excel

            path = Server.MapPath("ExcelMode/CopySheet2.xls");
            file = File.OpenRead(path);
            HSSFWorkbook workBook2 = new HSSFWorkbook(file);

            HSSFWorkbook workBook_new = new HSSFWorkbook();//新的Excel

            for (int i = 0; i < workBook1.NumberOfSheets; i++)
            {
                HSSFSheet sheet1 = workBook1.GetSheetAt(i) as HSSFSheet;
                sheet1.CopyTo(workBook_new, sheet1.SheetName, true, true);
            }
            //注意当sheet.name 为重复的时候 报错
            for (int i = 0; i < workBook2.NumberOfSheets; i++)
            {
                HSSFSheet sheet1 = workBook2.GetSheetAt(i) as HSSFSheet;

                HSSFSheet sheet_test = workBook_new.GetSheet(sheet1.SheetName) as HSSFSheet;
                string sheetName = "";
                if (sheet_test == null)
                {
                    //表示不存在
                    sheetName = sheet1.SheetName;
                }
                else
                {
                    //表示存在
                    sheetName = sheet1.SheetName + "_New";
                }
                //重新生成一个新的sheetName               
                sheet1.CopyTo(workBook_new, sheetName, true, true);
            }

            downLoadFile_HSSF(workBook_new, "Excel_2003_CopySheet.xls");
        }

        //带有DropDownList的Excel
        protected void lnkbtnExcel_2003_DropDownList_Click(object sender, EventArgs e)
        {
            HSSFWorkbook workBook = InitializeWorkbook_HSSF();
            ISheet sheet1 = workBook.CreateSheet("Sheet1");
            ISheet sheet2 = workBook.CreateSheet("Sheet2");
            //create three items in Sheet2
            IRow row0 = sheet2.CreateRow(0);
            ICell cell0 = row0.CreateCell(4);
            cell0.SetCellValue("Product1");

            row0 = sheet2.CreateRow(1);
            cell0 = row0.CreateCell(4);
            cell0.SetCellValue("Product2");

            row0 = sheet2.CreateRow(2);
            cell0 = row0.CreateCell(4);
            cell0.SetCellValue("Product3");

            CellRangeAddressList rangeList = new CellRangeAddressList();
            //add the data validation to the first column (1-100 rows) 
            rangeList.AddCellRangeAddress(new CellRangeAddress(0, 100, 0, 0));
            DVConstraint dvconstraint = DVConstraint.CreateFormulaListConstraint("Sheet2!$E1:$E3");
            HSSFDataValidation dataValidation = new
                    HSSFDataValidation(rangeList, dvconstraint);
            //add the data validation to sheet1
            ((HSSFSheet)sheet1).AddValidationData(dataValidation);
            downLoadFile_HSSF(workBook, "Excel_2003_DropDownList.xls");

            #region 其他的可以生产下拉框的代码
            // 1、
            //HSSFSheet sheet1 = hssfworkbook.CreateSheet("Sheet1");
            //CellRangeAddressList regions = new CellRangeAddressList(0, 65535, 0, 0);
            //DVConstraint constraint = DVConstraint.CreateExplicitListConstraint(new string[] { "itemA", "itemB", "itemC" });
            //HSSFDataValidation dataValidate = new HSSFDataValidation(regions, constraint);
            //sheet1.AddValidationData(dataValidate);
            // 2、

            //先创建一个Sheet专门用于存储下拉项的值，并将各下拉项的值写入其中：

            //HSSFSheet sheet2 = hssfworkbook.CreateSheet("ShtDictionary");
            //sheet2.CreateRow(0).CreateCell(0).SetCellValue("itemA");
            //sheet2.CreateRow(1).CreateCell(0).SetCellValue("itemB");
            //sheet2.CreateRow(2).CreateCell(0).SetCellValue("itemC");

            //然后定义一个名称，指向刚才创建的下拉项的区域：

            //HSSFName range = hssfworkbook.CreateName();
            //range.Reference = "ShtDictionary!$A1:$A3";
            //range.NameName = "dicRange";

            //最后，设置数据约束时指向这个名称而不是字符数组：

            //HSSFSheet sheet1 = hssfworkbook.CreateSheet("Sheet1");
            //CellRangeAddressList regions = new CellRangeAddressList(0, 65535, 0, 0);

            //DVConstraint constraint = DVConstraint.CreateFormulaListConstraint("dicRange");
            //HSSFDataValidation dataValidate = new HSSFDataValidation(regions, constraint);
            //sheet1.AddValidationData(dataValidate);

            //    3、

            // HSSFDataValidation   dataValidation = this.GetDataListValidation(result, colIndex);
            // sht.AddValidationData(dataValidation);

            // /**
            // * 设置某区域的有效性规则(列表)
            // * @return 生成的有效性规则
            // */
            //private HSSFDataValidation GetDataListValidation(string[] list, int colIndex)
            //{
            //    //设置数据有效性作用域
            //    CellRangeAddressList regions = GetRegionByColIndex(colIndex);

            //    //生成下拉框内容
            //    DVConstraint constraint = DVConstraint.CreateExplicitListConstraint(list);

            //    //绑定下拉框和作用区域
            //    HSSFDataValidation data_validation = new HSSFDataValidation(regions, constraint);
            //    //data_validation.CreateErrorBox("输入不合法", "请输入下拉列表中的值。");
            //    return data_validation;
            //}

            ////根据列序号获取整列区域
            //private CellRangeAddressList GetRegionByColIndex(int colIndex)
            //{
            //    return new CellRangeAddressList(1, 65535, colIndex, colIndex);
            //}
            #endregion
        }

        //带有页眉和页脚的Excel
        protected void lnkbtnExcel_2003_HeadAndFoot_Click(object sender, EventArgs e)
        {
            HSSFWorkbook workBook = InitializeWorkbook_HSSF();

            ISheet s1 = workBook.CreateSheet("Sheet1");
            s1.CreateRow(0).CreateCell(1).SetCellValue(123);

            //set header text
            s1.Header.Left = HSSFHeader.Page;   //Page is a static property of HSSFHeader and HSSFFooter
            s1.Header.Center = "This is a test sheet";
            //set footer text
            s1.Footer.Left = "Copyright NPOI Team";
            s1.Footer.Right = "created by 谢明昊";

            downLoadFile_HSSF(workBook, "Excel_2003_HeadAndFoot.xls");
        }

        ////带有自定义颜色的Excel
        protected void lnkbtnExcel_2003_CustomColor_Click(object sender, EventArgs e)
        {
            HSSFWorkbook workBook = InitializeWorkbook_HSSF();

            HSSFPalette palette = workBook.GetCustomPalette();
            palette.SetColorAtIndex(HSSFColor.Pink.Index, (byte)255, (byte)234, (byte)222);
            //HSSFColor myColor = palette.AddColor((byte)253, (byte)0, (byte)0);

            ISheet sheet1 = workBook.CreateSheet("Sheet1");
            ICellStyle style1 = workBook.CreateCellStyle();
            style1.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Pink.Index;
            style1.FillPattern = FillPattern.SolidForeground;
            sheet1.CreateRow(0).CreateCell(0).CellStyle = style1;

            downLoadFile_HSSF(workBook, "Excel_2003_CustomColor.xls");
        }

        //带有网格线的Excel
        protected void lnkbtnExcel_2003_ShowGridLines_Click(object sender, EventArgs e)
        {
            HSSFWorkbook workBook = InitializeWorkbook_HSSF();

            //sheet1 disables gridline
            ISheet s1 = workBook.CreateSheet("Sheet1");
            s1.DisplayGridlines = false;

            //sheet2 enables gridline
            ISheet s2 = workBook.CreateSheet("Sheet2");
            s2.DisplayGridlines = true;

            downLoadFile_HSSF(workBook, "Excel_2003_ShowGridLines.xls");
        }

        //带有绘画的Excel
        protected void lnkbtnExcel_2003_Drawing_Click(object sender, EventArgs e)
        {
            HSSFWorkbook workBook = InitializeWorkbook_HSSF();

            ISheet sheet1 = workBook.CreateSheet("new sheet");
            ISheet sheet2 = workBook.CreateSheet("second sheet");
            ISheet sheet3 = workBook.CreateSheet("third sheet");
            ISheet sheet4 = workBook.CreateSheet("fourth sheet");

            // Draw stuff in them
            DrawSheet1(sheet1);
            DrawSheet2(sheet2);
            DrawSheet3(sheet3);
            DrawSheet4(sheet4, workBook);

            downLoadFile_HSSF(workBook, "Excel_2003_Drawing.xls");
        }

        #region 带有绘画的Excel 的子方法

        private static void DrawSheet1(ISheet sheet1)
        {
            // Create a row and size one of the cells reasonably large.
            IRow row = sheet1.CreateRow(2);
            row.Height = ((short)2800);
            row.CreateCell(1);
            sheet1.SetColumnWidth(2, 9000);

            // Create the Drawing patriarch.  This is the top level container for
            // all shapes.
            HSSFPatriarch patriarch = (HSSFPatriarch)sheet1.CreateDrawingPatriarch();

            // Draw some lines and an oval.
            DrawLinesToCenter(patriarch);
            DrawManyLines(patriarch);
            DrawOval(patriarch);
            DrawPolygon(patriarch);

            // Draw a rectangle.
            HSSFSimpleShape rect = patriarch.CreateSimpleShape(new HSSFClientAnchor(100, 100, 900, 200, (short)0, 0, (short)0, 0));
            rect.ShapeType = (HSSFSimpleShape.OBJECT_TYPE_RECTANGLE);
        }

        private static void DrawSheet2(ISheet sheet2)
        {
            // Create a row and size one of the cells reasonably large.
            IRow row = sheet2.CreateRow(2);
            row.CreateCell(1);
            row.HeightInPoints = 240;
            sheet2.SetColumnWidth(2, 9000);

            // Create the Drawing patriarch.  This is the top level container for
            // all shapes. This will clear out any existing shapes for that sheet.
            HSSFPatriarch patriarch = (HSSFPatriarch)sheet2.CreateDrawingPatriarch();

            // Draw a grid in one of the cells.
            DrawGrid(patriarch);
        }

        private static void DrawSheet3(ISheet sheet3)
        {
            // Create a row and size one of the cells reasonably large
            IRow row = sheet3.CreateRow(2);
            row.HeightInPoints = 140;
            row.CreateCell(1);
            sheet3.SetColumnWidth(2, 9000);

            // Create the Drawing patriarch.  This is the top level container for
            // all shapes. This will clear out any existing shapes for that sheet.
            HSSFPatriarch patriarch = (HSSFPatriarch)sheet3.CreateDrawingPatriarch();

            // Create a shape group.
            HSSFShapeGroup group = patriarch.CreateGroup(
                    new HSSFClientAnchor(0, 0, 900, 200, (short)2, 2, (short)2, 2));

            // Create a couple of lines in the group.
            HSSFSimpleShape shape1 = group.CreateShape(new HSSFChildAnchor(3, 3, 500, 500));
            shape1.ShapeType = (HSSFSimpleShape.OBJECT_TYPE_LINE);
            ((HSSFChildAnchor)shape1.Anchor).SetAnchor((short)3, 3, 500, 500);
            HSSFSimpleShape shape2 = group.CreateShape(new HSSFChildAnchor((short)1, 200, 400, 600));
            shape2.ShapeType = (HSSFSimpleShape.OBJECT_TYPE_LINE);

        }

        private static void DrawSheet4(ISheet sheet4, HSSFWorkbook wb)
        {
            // Create the Drawing patriarch.  This is the top level container for
            // all shapes. This will clear out any existing shapes for that sheet.
            HSSFPatriarch patriarch = (HSSFPatriarch)sheet4.CreateDrawingPatriarch();

            // Create a couple of textboxes
            HSSFTextbox textbox1 = (HSSFTextbox)patriarch.CreateTextbox(
                    new HSSFClientAnchor(0, 0, 0, 0, (short)1, 1, (short)2, 2));
            textbox1.String = new HSSFRichTextString("This is a test");
            HSSFTextbox textbox2 = (HSSFTextbox)patriarch.CreateTextbox(
                    new HSSFClientAnchor(0, 0, 900, 100, (short)3, 3, (short)3, 4));
            textbox2.String = new HSSFRichTextString("Woo");
            textbox2.SetFillColor(200, 0, 0);
            textbox2.LineStyle = LineStyle.DotGel;

            // Create third one with some fancy font styling.
            HSSFTextbox textbox3 = (HSSFTextbox)patriarch.CreateTextbox(
                    new HSSFClientAnchor(0, 0, 900, 100, (short)4, 4, (short)5, 4 + 1));
            IFont font = wb.CreateFont();
            font.IsItalic = true;
            font.Underline = FontUnderlineType.Double;
            HSSFRichTextString str = new HSSFRichTextString("Woo!!!");
            str.ApplyFont(2, 5, font);
            textbox3.String = str;
            textbox3.FillColor = 0x08000030;
            textbox3.LineStyle = LineStyle.None;  // no line around the textbox.
            textbox3.IsNoFill = true;    // make it transparent
        }

        private static void DrawOval(HSSFPatriarch patriarch)
        {
            // Create an oval and style to taste.
            HSSFClientAnchor a = new HSSFClientAnchor();
            a.SetAnchor((short)2, 2, 20, 20, (short)2, 2, 190, 80);
            HSSFSimpleShape s = patriarch.CreateSimpleShape(a);
            s.ShapeType = HSSFSimpleShape.OBJECT_TYPE_OVAL;
            s.SetLineStyleColor(10, 10, 10);
            s.SetFillColor(90, 10, 200);
            s.LineWidth = HSSFShape.LINEWIDTH_ONE_PT * 3;
            s.LineStyle = LineStyle.DotSys;
        }

        private static void DrawPolygon(HSSFPatriarch patriarch)
        {
            HSSFClientAnchor a = new HSSFClientAnchor();
            a.SetAnchor((short)2, 2, 0, 0, (short)3, 3, 1023, 255);
            HSSFShapeGroup g = patriarch.CreateGroup(a);
            g.SetCoordinates(0, 0, 200, 200);
            HSSFPolygon p1 = g.CreatePolygon(new HSSFChildAnchor(0, 0, 200, 200));
            p1.SetPolygonDrawArea(100, 100);
            p1.SetPoints(new int[] { 0, 90, 50 }, new int[] { 5, 5, 44 });
            p1.SetFillColor(0, 255, 0);
            HSSFPolygon p2 = g.CreatePolygon(new HSSFChildAnchor(20, 20, 200, 200));
            p2.SetPolygonDrawArea(200, 200);
            p2.SetPoints(new int[] { 120, 20, 150 }, new int[] { 105, 30, 195 });
            p2.SetFillColor(255, 0, 0);
        }

        private static void DrawManyLines(HSSFPatriarch patriarch)
        {
            // Draw bunch of lines
            int x1 = 100;
            int y1 = 100;
            int x2 = 800;
            int y2 = 200;
            int color = 0;
            for (int i = 0; i < 10; i++)
            {
                HSSFClientAnchor a2 = new HSSFClientAnchor();
                a2.SetAnchor((short)2, 2, x1, y1, (short)2, 2, x2, y2);
                HSSFSimpleShape shape2 = patriarch.CreateSimpleShape(a2);
                shape2.ShapeType = HSSFSimpleShape.OBJECT_TYPE_LINE;
                shape2.LineStyleColor = color;
                y1 -= 10;
                y2 -= 10;
                color += 30;
            }
        }

        private static void DrawGrid(HSSFPatriarch patriarch)
        {
            // This Draws a grid of lines.  Since the coordinates space fixed at
            // 1024 by 256 we use a ratio to get a reasonably square grids.

            double xRatio = 3.22;
            double yRatio = 0.6711;

            int x1 = 000;
            int y1 = 000;
            int x2 = 000;
            int y2 = 200;
            for (int i = 0; i < 20; i++)
            {
                HSSFClientAnchor a2 = new HSSFClientAnchor();
                a2.SetAnchor((short)2, 2, (int)(x1 * xRatio), (int)(y1 * yRatio),
                        (short)2, 2, (int)(x2 * xRatio), (int)(y2 * yRatio));
                HSSFSimpleShape shape2 = patriarch.CreateSimpleShape(a2);
                shape2.ShapeType = (HSSFSimpleShape.OBJECT_TYPE_LINE);

                x1 += 10;
                x2 += 10;
            }

            x1 = 000;
            y1 = 000;
            x2 = 200;
            y2 = 000;
            for (int i = 0; i < 20; i++)
            {
                HSSFClientAnchor a2 = new HSSFClientAnchor();
                a2.SetAnchor((short)2, 2, (int)(x1 * xRatio), (int)(y1 * yRatio),
                        (short)2, 2, (int)(x2 * xRatio), (int)(y2 * yRatio));
                HSSFSimpleShape shape2 = patriarch.CreateSimpleShape(a2);
                shape2.ShapeType = (HSSFSimpleShape.OBJECT_TYPE_LINE);

                y1 += 10;
                y2 += 10;
            }
        }

        private static void DrawLinesToCenter(HSSFPatriarch patriarch)
        {
            // Draw some lines from and to the corners
            {
                HSSFClientAnchor a1 = new HSSFClientAnchor();
                a1.SetAnchor((short)2, 2, 0, 0, (short)2, 2, 512, 128);
                HSSFSimpleShape shape1 = patriarch.CreateSimpleShape(a1);
                shape1.ShapeType = (HSSFSimpleShape.OBJECT_TYPE_LINE);
            }
            {
                HSSFClientAnchor a1 = new HSSFClientAnchor();
                a1.SetAnchor((short)2, 2, 512, 128, (short)2, 2, 1023, 0);
                HSSFSimpleShape shape1 = patriarch.CreateSimpleShape(a1);
                shape1.ShapeType = (HSSFSimpleShape.OBJECT_TYPE_LINE);
            }
            {
                HSSFClientAnchor a1 = new HSSFClientAnchor();
                a1.SetAnchor((short)1, 1, 0, 0, (short)1, 1, 512, 100);
                HSSFSimpleShape shape1 = patriarch.CreateSimpleShape(a1);
                shape1.ShapeType = (HSSFSimpleShape.OBJECT_TYPE_LINE);
            }
            {
                HSSFClientAnchor a1 = new HSSFClientAnchor();
                a1.SetAnchor((short)1, 1, 512, 100, (short)1, 1, 1023, 0);
                HSSFSimpleShape shape1 = patriarch.CreateSimpleShape(a1);
                shape1.ShapeType = (HSSFSimpleShape.OBJECT_TYPE_LINE);
            }

        }

        #endregion


        //带有排序筛选的Excel
        protected void lnkbtnExcel_2003_AutoFilter_Click(object sender, EventArgs e)
        {
            HSSFWorkbook workBook = InitializeWorkbook_HSSF();

            HSSFSheet sheet1 = (HSSFSheet)workBook.CreateSheet("Sheet1");
            //create horizontal 1-9
            for (int i = 1; i <= 9; i++)
            {
                IRow row = sheet1.CreateRow(i);
                //create vertical 1-9
                for (int j = 1; j <= 9; j++)
                {
                    row.CreateCell(j).SetCellValue(i * j);
                }
            }

            sheet1.SetAutoFilter(new CellRangeAddress(1, 9, 1, 5));

            downLoadFile_HSSF(workBook, "Excel_2003_AutoFilter.xls");
        }

        //提取Excel中的图片
        protected void lnkbtnExcel_2003_ExtractPictures_Click(object sender, EventArgs e)
        {
            string path = Server.MapPath("ExcelMode/Extract.xls");
            FileStream file = File.OpenRead(path);
            HSSFWorkbook workBook = new HSSFWorkbook(file);

            System.Collections.IList pictures = workBook.GetAllPictures();
            int i = 0;
            foreach (HSSFPictureData pic in pictures)
            {
                string ext = pic.SuggestFileExtension();
                if (ext.Equals("jpeg"))
                {
                    System.Drawing.Image jpg = System.Drawing.Image.FromStream(new MemoryStream(pic.Data));
                    jpg.Save(string.Format(@"D:\pic{0}.jpg", i++));
                }
                else if (ext.Equals("png"))
                {
                    System.Drawing.Image png = System.Drawing.Image.FromStream(new MemoryStream(pic.Data));
                    png.Save(string.Format(@"D:\pic{0}.png", i++));
                }
            }

        }

        //提取Excel中的字符串
        protected void lnkbtnExcel_2003_ExtractString_Click(object sender, EventArgs e)
        {
            string path = Server.MapPath("ExcelMode/Extract.xls");

            Stream file = new FileStream(path, FileMode.Open);
            MemoryStream ms = new MemoryStream();
            byte[] buf = new byte[512];
            while (true)
            {
                int bytesRead = file.Read(buf, 0, buf.Length);
                if (bytesRead < 1)
                {
                    break;
                }
                ms.Write(buf, 0, bytesRead);
            }
            file.Close();

            //设置响应的类型为Excel
            Response.ContentType = "application/vnd.ms-excel";
            //设置下载的Excel文件名
            Response.AddHeader("Content-Disposition", string.Format("attachment;filename={0}", "Excel_2003_ExtractString.xls"));//目前只能为03的Excel
            //Clear方法删除所有缓存中的HTML输出。但此方法只删除Response显示输入信息，不删除Response头信息。以免影响导出数据的完整性。
            Response.Clear();
            ms.WriteTo(Response.OutputStream);
        }

        //带有背景填充的Excel
        protected void lnkbtnExcel_2003_FillBackground_Click(object sender, EventArgs e)
        {
            HSSFWorkbook workBook = InitializeWorkbook_HSSF();
            ISheet sheet1 = workBook.CreateSheet("Sheet1");

            //fill background
            ICellStyle style1 = workBook.CreateCellStyle();
            style1.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Blue.Index;
            style1.FillPattern = FillPattern.BigSpots;
            style1.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.Pink.Index;
            sheet1.CreateRow(0).CreateCell(0).CellStyle = style1;

            //fill background
            ICellStyle style2 = workBook.CreateCellStyle();
            style2.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Yellow.Index;
            style2.FillPattern = FillPattern.AltBars;
            style2.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.Rose.Index;
            sheet1.CreateRow(1).CreateCell(0).CellStyle = style2;

            //fill background
            ICellStyle style3 = workBook.CreateCellStyle();
            style3.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Lime.Index;
            style3.FillPattern = FillPattern.LessDots;
            style3.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.LightGreen.Index;
            sheet1.CreateRow(2).CreateCell(0).CellStyle = style3;

            //fill backgroundworkBook.CreateCellStyle();
            ICellStyle style4 = workBook.CreateCellStyle();
            style4.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Yellow.Index;
            style4.FillPattern = FillPattern.LeastDots;
            style4.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.Rose.Index;
            sheet1.CreateRow(3).CreateCell(0).CellStyle = style4;

            //fill background
            ICellStyle style5 = workBook.CreateCellStyle();
            style5.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.LightBlue.Index;
            style5.FillPattern = FillPattern.Bricks;
            style5.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.Plum.Index;
            sheet1.CreateRow(4).CreateCell(0).CellStyle = style5;

            //fill background
            ICellStyle style6 = workBook.CreateCellStyle();
            style6.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.SeaGreen.Index;
            style6.FillPattern = FillPattern.FineDots;
            style6.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.White.Index;
            sheet1.CreateRow(5).CreateCell(0).CellStyle = style6;

            //fill background
            ICellStyle style7 = workBook.CreateCellStyle();
            style7.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Orange.Index;
            style7.FillPattern = FillPattern.Diamonds;
            style7.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.Orchid.Index;
            sheet1.CreateRow(6).CreateCell(0).CellStyle = style7;

            //fill background
            ICellStyle style8 = workBook.CreateCellStyle();
            style8.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.White.Index;
            style8.FillPattern = FillPattern.Squares;
            style8.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.Red.Index;
            sheet1.CreateRow(7).CreateCell(0).CellStyle = style8;

            //fill background
            ICellStyle style9 = workBook.CreateCellStyle();
            style9.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.RoyalBlue.Index;
            style9.FillPattern = FillPattern.SparseDots;
            style9.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.Yellow.Index;
            sheet1.CreateRow(8).CreateCell(0).CellStyle = style9;

            //fill background
            ICellStyle style10 = workBook.CreateCellStyle();
            style10.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.RoyalBlue.Index;
            style10.FillPattern = FillPattern.ThickBackwardDiagonals;
            style10.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.Yellow.Index;
            sheet1.CreateRow(9).CreateCell(0).CellStyle = style10;

            //fill background
            ICellStyle style11 = workBook.CreateCellStyle();
            style11.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.RoyalBlue.Index;
            style11.FillPattern = FillPattern.ThickForwardDiagonals;
            style11.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.Yellow.Index;
            sheet1.CreateRow(10).CreateCell(0).CellStyle = style11;

            //fill background
            ICellStyle style12 = workBook.CreateCellStyle();
            style12.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.RoyalBlue.Index;
            style12.FillPattern = FillPattern.ThickHorizontalBands;
            style12.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.Yellow.Index;
            sheet1.CreateRow(11).CreateCell(0).CellStyle = style12;


            //fill background
            ICellStyle style13 = workBook.CreateCellStyle();
            style13.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.RoyalBlue.Index;
            style13.FillPattern = FillPattern.ThickVerticalBands;
            style13.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.Yellow.Index;
            sheet1.CreateRow(12).CreateCell(0).CellStyle = style13;

            //fill background
            ICellStyle style14 = workBook.CreateCellStyle();
            style14.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.RoyalBlue.Index;
            style14.FillPattern = FillPattern.ThinBackwardDiagonals;
            style14.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.Yellow.Index;
            sheet1.CreateRow(13).CreateCell(0).CellStyle = style14;

            //fill background
            ICellStyle style15 = workBook.CreateCellStyle();
            style15.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.RoyalBlue.Index;
            style15.FillPattern = FillPattern.ThinForwardDiagonals;
            style15.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.Yellow.Index;
            sheet1.CreateRow(14).CreateCell(0).CellStyle = style15;

            //fill background
            ICellStyle style16 = workBook.CreateCellStyle();
            style16.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.RoyalBlue.Index;
            style16.FillPattern = FillPattern.ThinHorizontalBands;
            style16.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.Yellow.Index;
            sheet1.CreateRow(15).CreateCell(0).CellStyle = style16;

            //fill background
            ICellStyle style17 = workBook.CreateCellStyle();
            style17.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.RoyalBlue.Index;
            style17.FillPattern = FillPattern.ThinVerticalBands;
            style17.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.Yellow.Index;
            sheet1.CreateRow(16).CreateCell(0).CellStyle = style17;

            downLoadFile_HSSF(workBook, "Excel_2003_FillBackground.xls");
        }

        //读取模板生成Excel
        protected void lnkbtnExcel_2003_GenerateFromTemplate_Click(object sender, EventArgs e)
        {
            FileStream file = new FileStream(Server.MapPath("ExcelMode/Template.xls"), FileMode.Open, FileAccess.Read);

            HSSFWorkbook workBook = new HSSFWorkbook(file);

            ISheet sheet1 = workBook.GetSheet("Sheet1");
            //create cell on rows, since rows do already exist,it's not necessary to create rows again.
            sheet1.GetRow(1).GetCell(1).SetCellValue(200200);
            sheet1.GetRow(2).GetCell(1).SetCellValue(300);
            sheet1.GetRow(3).GetCell(1).SetCellValue(500050);
            sheet1.GetRow(4).GetCell(1).SetCellValue(8000);
            sheet1.GetRow(5).GetCell(1).SetCellValue(110);
            sheet1.GetRow(6).GetCell(1).SetCellValue(100);
            sheet1.GetRow(7).GetCell(1).SetCellValue(200);
            sheet1.GetRow(8).GetCell(1).SetCellValue(210);
            sheet1.GetRow(9).GetCell(1).SetCellValue(2300);
            sheet1.GetRow(10).GetCell(1).SetCellValue(240);
            sheet1.GetRow(11).GetCell(1).SetCellValue(180123);
            sheet1.GetRow(12).GetCell(1).SetCellValue(150);

            //Force excel to recalculate all the formula while open
            sheet1.ForceFormulaRecalculation = true;

            downLoadFile_HSSF(workBook, "Excel_2003_GenerateFromTemplate.xls");
        }

        protected void lnkbtnExcel_2003_FromTemplateChart_Click(object sender, EventArgs e)
        {
            //柱状图 就只好先做个模板 然后在导入数据   记得模板中再添加数据透视表/图后 选择数据中自动跟新/打开时候更新 要不然没反应

            string path = Server.MapPath("ExcelMode/chart_Model.xls");

            HSSFWorkbook workBook = new HSSFWorkbook(new FileStream(path, FileMode.Open));
            //注意是GetSheet
            ISheet sheet = workBook.GetSheet("Sheet1");

            //IRow row = sheet.CreateRow(0);
            //row.CreateCell(0).SetCellValue("姓名");
            //row.CreateCell(1).SetCellValue("销售额");

            IRow row = sheet.CreateRow(1);
            row.CreateCell(0).SetCellValue("令狐冲");
            row.CreateCell(1).SetCellValue(50000);

            row = sheet.CreateRow(2);
            row.CreateCell(0).SetCellValue("任盈盈");
            row.CreateCell(1).SetCellValue(30000);

            row = sheet.CreateRow(3);
            row.CreateCell(0).SetCellValue("风清扬");
            row.CreateCell(1).SetCellValue(80000);

            row = sheet.CreateRow(4);
            row.CreateCell(0).SetCellValue("任我行");
            row.CreateCell(1).SetCellValue(20000);

            row = sheet.CreateRow(5);
            row.CreateCell(0).SetCellValue("左冷禅");
            row.CreateCell(1).SetCellValue(10000);

            downLoadFile_HSSF(workBook, "Excel_2003_FromTemplateChart.xls");
        }

        //带有组(行,列)的Excel
        protected void lnkbtnExcel_2003_GroupRowAndColumn_Click(object sender, EventArgs e)
        {
            HSSFWorkbook workBook = InitializeWorkbook_HSSF();
            ISheet s = workBook.CreateSheet("Sheet1");
            IRow r1 = s.CreateRow(0);
            IRow r2 = s.CreateRow(1);
            IRow r3 = s.CreateRow(2);
            IRow r4 = s.CreateRow(3);
            IRow r5 = s.CreateRow(4);

            //group row 2 to row 4
            s.GroupRow(1, 3);

            //group row 2 to row 3
            s.GroupRow(1, 2);

            //group column 1-3
            s.GroupColumn(1, 3);
            downLoadFile_HSSF(workBook, "Excel_2003_GroupRowAndColumn.xls");
        }

        //带有隐藏行和列的Excel
        protected void lnkbtnExcel_2003_HideRowAndColumn_Click(object sender, EventArgs e)
        {
            HSSFWorkbook workBook = InitializeWorkbook_HSSF();

            ISheet s = workBook.CreateSheet("Sheet1");
            IRow r1 = s.CreateRow(0);
            IRow r2 = s.CreateRow(1);
            IRow r3 = s.CreateRow(2);
            IRow r4 = s.CreateRow(3);
            IRow r5 = s.CreateRow(4);

            //hide IRow 2
            r2.ZeroHeight = true;

            //hide column C
            s.SetColumnHidden(2, true);

            downLoadFile_HSSF(workBook, "Excel_2003_HideRowAndColumn.xls");
        }

        //导入Excel数据,转换成dataTable在读取导出
        protected void lnkbtnExcel_2003_ImportExcel_Click(object sender, EventArgs e)
        {
            FileStream file = new FileStream(Server.MapPath("ExcelMode/ImportExcel.xls"), FileMode.Open);
            HSSFWorkbook workBook = new HSSFWorkbook(file);

            ISheet sheet = workBook.GetSheetAt(0);
            System.Collections.IEnumerator rows = sheet.GetRowEnumerator();

            DataTable dt = new DataTable();
            for (int j = 0; j < 5; j++)
            {
                dt.Columns.Add(Convert.ToChar(((int)'A') + j).ToString());
            }

            while (rows.MoveNext())
            {
                IRow row = (HSSFRow)rows.Current;
                DataRow dr = dt.NewRow();

                for (int i = 0; i < row.LastCellNum; i++)
                {
                    ICell cell = row.GetCell(i);


                    if (cell == null)
                    {
                        dr[i] = null;
                    }
                    else
                    {
                        dr[i] = cell.ToString();
                    }
                }
                dt.Rows.Add(dr);
            }


            HSSFWorkbook workBook_new = InitializeWorkbook_HSSF();
            ISheet sheet_new = workBook_new.CreateSheet("aaa");
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                IRow row = sheet_new.CreateRow(i);
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    ICell cell = row.CreateCell(j);
                    if (dt.Rows[i][j] == null || dt.Rows[i][j].ToString() == "")
                    {
                        // cell.SetCellValue("0.0");
                    }
                    else
                    {
                        cell.SetCellValue(dt.Rows[i][j].ToString());
                    }
                }
            }

            downLoadFile_HSSF(workBook_new, "Excel_2003_ImportExcel.xls");
        }

        //带有图片的Excel
        protected void lnkbtnExcel_2003_Image_Click(object sender, EventArgs e)
        {
            HSSFWorkbook workBook = InitializeWorkbook_HSSF();
            ISheet sheet1 = workBook.CreateSheet("PictureSheet");


            HSSFPatriarch patriarch = (HSSFPatriarch)sheet1.CreateDrawingPatriarch();
            //create the anchor
            HSSFClientAnchor anchor;
            anchor = new HSSFClientAnchor(500, 200, 0, 0, 2, 2, 4, 7);
            anchor.AnchorType = 2;

            FileStream file = new FileStream(Server.MapPath("Images/HumpbackWhale.jpg"), FileMode.Open, FileAccess.Read);
            byte[] buffer = new byte[file.Length];
            file.Read(buffer, 0, (int)file.Length);
            int intImage = workBook.AddPicture(buffer, NPOI.SS.UserModel.PictureType.JPEG);

            //load the picture and get the picture index in the workbook
            HSSFPicture picture = (HSSFPicture)patriarch.CreatePicture(anchor, intImage);
            //Reset the image to the original size.
            //picture.Resize();   //Note: Resize will reset client anchor you set.
            picture.LineStyle = LineStyle.DashDotGel;

            downLoadFile_HSSF(workBook, "Excel_2003_Image.xls");
        }

        #region 带有贷款计算器的Excel

        protected void lnkbtnExcel_2003_LoanCalculator_Click(object sender, EventArgs e)
        {
            HSSFWorkbook workBook = InitializeWorkbook_HSSF();

            Dictionary<String, ICellStyle> styles = CreateStyles(workBook);
            ISheet sheet = workBook.CreateSheet("Loan Calculator");
            sheet.IsPrintGridlines = (false);
            sheet.DisplayGridlines = (false);

            IPrintSetup printSetup = sheet.PrintSetup;
            printSetup.Landscape = (true);
            sheet.FitToPage = (true);
            sheet.HorizontallyCenter = (true);

            sheet.SetColumnWidth(0, 3 * 256);
            sheet.SetColumnWidth(1, 3 * 256);
            sheet.SetColumnWidth(2, 11 * 256);
            sheet.SetColumnWidth(3, 14 * 256);
            sheet.SetColumnWidth(4, 14 * 256);
            sheet.SetColumnWidth(5, 14 * 256);
            sheet.SetColumnWidth(6, 14 * 256);

            IName name;

            name = workBook.CreateName();
            name.NameName = ("Interest_Rate");
            name.RefersToFormula = ("'Loan Calculator'!$E$5");

            name = workBook.CreateName();
            name.NameName = ("Loan_Amount");
            name.RefersToFormula = ("'Loan Calculator'!$E$4");

            name = workBook.CreateName();
            name.NameName = ("Loan_Start");
            name.RefersToFormula = ("'Loan Calculator'!$E$7");

            name = workBook.CreateName();
            name.NameName = ("Loan_Years");
            name.RefersToFormula = ("'Loan Calculator'!$E$6");

            name = workBook.CreateName();
            name.NameName = ("Number_of_Payments");
            name.RefersToFormula = ("'Loan Calculator'!$E$10");

            name = workBook.CreateName();
            name.NameName = ("Monthly_Payment");
            name.RefersToFormula = ("-PMT(Interest_Rate/12,Number_of_Payments,Loan_Amount)");

            name = workBook.CreateName();
            name.NameName = ("Total_Cost");
            name.RefersToFormula = ("'Loan Calculator'!$E$12");

            name = workBook.CreateName();
            name.NameName = ("Total_Interest");
            name.RefersToFormula = ("'Loan Calculator'!$E$11");

            name = workBook.CreateName();
            name.NameName = ("Values_Entered");
            name.RefersToFormula = ("IF(ISBLANK(Loan_Start),0,IF(Loan_Amount*Interest_Rate*Loan_Years>0,1,0))");

            IRow titleRow = sheet.CreateRow(0);
            titleRow.HeightInPoints = (35);
            for (int i = 1; i <= 7; i++)
            {
                titleRow.CreateCell(i).CellStyle = styles["title"];
            }
            ICell titleCell = titleRow.GetCell(2);
            titleCell.SetCellValue("Simple Loan Calculator");
            sheet.AddMergedRegion(CellRangeAddress.ValueOf("$C$1:$H$1"));

            IRow row = sheet.CreateRow(2);
            ICell cell = row.CreateCell(4);
            cell.SetCellValue("Enter values");
            cell.CellStyle = styles["item_right"];

            row = sheet.CreateRow(3);
            cell = row.CreateCell(2);
            cell.SetCellValue("Loan amount");
            cell.CellStyle = styles["item_left"];
            cell = row.CreateCell(4);
            cell.CellStyle = styles["input_$"];
            cell.SetAsActiveCell();

            row = sheet.CreateRow(4);
            cell = row.CreateCell(2);
            cell.SetCellValue("Annual interest rate");
            cell.CellStyle = styles["item_left"];
            cell = row.CreateCell(4);
            cell.CellStyle = styles["input_%"];

            row = sheet.CreateRow(5);
            cell = row.CreateCell(2);
            cell.SetCellValue("Loan period in years");
            cell.CellStyle = styles["item_left"];
            cell = row.CreateCell(4);
            cell.CellStyle = styles["input_i"];

            row = sheet.CreateRow(6);
            cell = row.CreateCell(2);
            cell.SetCellValue("Start date of loan");
            cell.CellStyle = styles["item_left"];
            cell = row.CreateCell(4);
            cell.CellStyle = styles["input_d"];

            row = sheet.CreateRow(8);
            cell = row.CreateCell(2);
            cell.SetCellValue("Monthly payment");
            cell.CellStyle = styles["item_left"];
            cell = row.CreateCell(4);
            cell.CellFormula = ("IF(Values_Entered,Monthly_Payment,\"\")");
            cell.CellStyle = styles["formula_$"];

            row = sheet.CreateRow(9);
            cell = row.CreateCell(2);
            cell.SetCellValue("Number of payments");
            cell.CellStyle = styles["item_left"];
            cell = row.CreateCell(4);
            cell.CellFormula = ("IF(Values_Entered,Loan_Years*12,\"\")");
            cell.CellStyle = styles["formula_i"];

            row = sheet.CreateRow(10);
            cell = row.CreateCell(2);
            cell.SetCellValue("Total interest");
            cell.CellStyle = styles["item_left"];
            cell = row.CreateCell(4);
            cell.CellFormula = ("IF(Values_Entered,Total_Cost-Loan_Amount,\"\")");
            cell.CellStyle = styles["formula_$"];

            row = sheet.CreateRow(11);
            cell = row.CreateCell(2);
            cell.SetCellValue("Total cost of loan");
            cell.CellStyle = styles["item_left"];
            cell = row.CreateCell(4);
            cell.CellFormula = ("IF(Values_Entered,Monthly_Payment*Number_of_Payments,\"\")");
            cell.CellStyle = styles["formula_$"];

            downLoadFile_HSSF(workBook, "Excel_2003_LoanCalculator.xls");
        }

        //创建贷款计算器的样式
        private Dictionary<String, ICellStyle> CreateStyles(IWorkbook wb)
        {
            Dictionary<String, ICellStyle> styles = new Dictionary<String, ICellStyle>();

            ICellStyle style = null;
            IFont titleFont = wb.CreateFont();
            titleFont.FontHeightInPoints = (short)14;
            titleFont.FontName = "Trebuchet MS";
            style = wb.CreateCellStyle();
            style.SetFont(titleFont);
            style.BorderBottom = NPOI.SS.UserModel.BorderStyle.Dotted;
            style.BottomBorderColor = IndexedColors.Grey40Percent.Index;
            styles.Add("title", style);

            IFont itemFont = wb.CreateFont();
            itemFont.FontHeightInPoints = (short)9;
            itemFont.FontName = "Trebuchet MS";
            style = wb.CreateCellStyle();
            style.Alignment = (HorizontalAlignment.Left);
            style.SetFont(itemFont);
            styles.Add("item_left", style);

            style = wb.CreateCellStyle();
            style.Alignment = HorizontalAlignment.Right;
            style.SetFont(itemFont);
            styles.Add("item_right", style);

            style = wb.CreateCellStyle();
            style.Alignment = HorizontalAlignment.Right;
            style.SetFont(itemFont);
            style.BorderRight = NPOI.SS.UserModel.BorderStyle.Dotted;
            style.RightBorderColor = IndexedColors.Grey40Percent.Index;
            style.BorderBottom = NPOI.SS.UserModel.BorderStyle.Dotted;
            style.BottomBorderColor = IndexedColors.Grey40Percent.Index;
            style.BorderLeft = NPOI.SS.UserModel.BorderStyle.Dotted;
            style.LeftBorderColor = IndexedColors.Grey40Percent.Index;
            style.BorderTop = NPOI.SS.UserModel.BorderStyle.Dotted;
            style.TopBorderColor = IndexedColors.Grey40Percent.Index;
            style.DataFormat = (wb.CreateDataFormat().GetFormat("_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(@_)"));
            styles.Add("input_$", style);

            style = wb.CreateCellStyle();
            style.Alignment = HorizontalAlignment.Right;
            style.SetFont(itemFont);
            style.BorderRight = NPOI.SS.UserModel.BorderStyle.Dotted;
            style.RightBorderColor = IndexedColors.Grey40Percent.Index;
            style.BorderBottom = NPOI.SS.UserModel.BorderStyle.Dotted;
            style.BottomBorderColor = IndexedColors.Grey40Percent.Index;
            style.BorderLeft = NPOI.SS.UserModel.BorderStyle.Dotted;
            style.LeftBorderColor = IndexedColors.Grey40Percent.Index;
            style.BorderTop = NPOI.SS.UserModel.BorderStyle.Dotted;
            style.TopBorderColor = IndexedColors.Grey40Percent.Index;
            style.DataFormat = (wb.CreateDataFormat().GetFormat("0.000%"));
            styles.Add("input_%", style);

            style = wb.CreateCellStyle();
            style.Alignment = HorizontalAlignment.Right;
            style.SetFont(itemFont);
            style.BorderRight = NPOI.SS.UserModel.BorderStyle.Dotted;
            style.RightBorderColor = IndexedColors.Grey40Percent.Index;
            style.BorderBottom = NPOI.SS.UserModel.BorderStyle.Dotted;
            style.BottomBorderColor = IndexedColors.Grey40Percent.Index;
            style.BorderLeft = NPOI.SS.UserModel.BorderStyle.Dotted;
            style.LeftBorderColor = IndexedColors.Grey40Percent.Index;
            style.BorderTop = NPOI.SS.UserModel.BorderStyle.Dotted;
            style.TopBorderColor = IndexedColors.Grey40Percent.Index;
            style.DataFormat = wb.CreateDataFormat().GetFormat("0");
            styles.Add("input_i", style);

            style = wb.CreateCellStyle();
            style.Alignment = (HorizontalAlignment.Center);
            style.SetFont(itemFont);
            style.DataFormat = wb.CreateDataFormat().GetFormat("m/d/yy");
            styles.Add("input_d", style);

            style = wb.CreateCellStyle();
            style.Alignment = HorizontalAlignment.Right;
            style.SetFont(itemFont);
            style.BorderRight = NPOI.SS.UserModel.BorderStyle.Dotted;
            style.RightBorderColor = IndexedColors.Grey40Percent.Index;
            style.BorderBottom = NPOI.SS.UserModel.BorderStyle.Dotted;
            style.BottomBorderColor = IndexedColors.Grey40Percent.Index;
            style.BorderLeft = NPOI.SS.UserModel.BorderStyle.Dotted;
            style.LeftBorderColor = IndexedColors.Grey40Percent.Index;
            style.BorderTop = NPOI.SS.UserModel.BorderStyle.Dotted;
            style.TopBorderColor = IndexedColors.Grey40Percent.Index;
            style.DataFormat = wb.CreateDataFormat().GetFormat("$##,##0.00");
            style.BorderBottom = NPOI.SS.UserModel.BorderStyle.Dotted;
            style.BottomBorderColor = IndexedColors.Grey40Percent.Index;
            style.FillForegroundColor = IndexedColors.Grey25Percent.Index;
            style.FillPattern = FillPattern.SolidForeground;
            styles.Add("formula_$", style);

            style = wb.CreateCellStyle();
            style.Alignment = HorizontalAlignment.Right;
            style.SetFont(itemFont);
            style.BorderRight = NPOI.SS.UserModel.BorderStyle.Dotted;
            style.RightBorderColor = IndexedColors.Grey40Percent.Index;
            style.BorderBottom = NPOI.SS.UserModel.BorderStyle.Dotted;
            style.BottomBorderColor = IndexedColors.Grey40Percent.Index;
            style.BorderLeft = NPOI.SS.UserModel.BorderStyle.Dotted;
            style.LeftBorderColor = IndexedColors.Grey40Percent.Index;
            style.BorderTop = NPOI.SS.UserModel.BorderStyle.Dotted;
            style.TopBorderColor = IndexedColors.Grey40Percent.Index;
            style.DataFormat = wb.CreateDataFormat().GetFormat("0");
            style.BorderBottom = NPOI.SS.UserModel.BorderStyle.Dotted;
            style.BottomBorderColor = IndexedColors.Grey40Percent.Index;
            style.FillForegroundColor = IndexedColors.Grey25Percent.Index;
            style.FillPattern = (FillPattern.SolidForeground);
            styles.Add("formula_i", style);

            return styles;
        }

        #endregion

        //带有合并单元格的Excel
        protected void lnkbtnExcel_2003_MergeCells_Click(object sender, EventArgs e)
        {
            HSSFWorkbook workBook = InitializeWorkbook_HSSF();
            ISheet sheet = workBook.CreateSheet("new sheet");

            IRow row = sheet.CreateRow(0);
            row.HeightInPoints = 30;

            ICell cell = row.CreateCell(0);
            //set the title of the sheet
            cell.SetCellValue("Sales Report");

            ICellStyle style = workBook.CreateCellStyle();
            style.Alignment = HorizontalAlignment.Center;
            //create a font style
            IFont font = workBook.CreateFont();
            font.FontHeight = 20 * 20;
            style.SetFont(font);
            cell.CellStyle = style;

            //merged cells on single row
            //ATTENTION: don't use Region class, which is obsolete
            sheet.AddMergedRegion(new CellRangeAddress(0, 0, 0, 5));

            //merged cells on mutiple rows
            CellRangeAddress region = new CellRangeAddress(2, 4, 2, 4);
            sheet.AddMergedRegion(region);

            //set enclosed border for the merged region
            ((HSSFSheet)sheet).SetEnclosedBorderOfRegion(region, NPOI.SS.UserModel.BorderStyle.Dotted, NPOI.HSSF.Util.HSSFColor.Red.Index);
            downLoadFile_HSSF(workBook, "Excel_2003_MergeCells.xls");
        }


        #region 带有折叠效果的Excel
        protected void lnkbtnExcel_2003_Plication_Click(object sender, EventArgs e)
        {
            HSSFWorkbook workBook = InitializeWorkbook_HSSF();
            ISheet sheet1 = workBook.CreateSheet("Multiple Table");

            //create horizontal 1-9
            for (int i = 1; i <= 9; i++)
            {
                sheet1.CreateRow(0).CreateCell(i).SetCellValue(i);
            }
            //create vertical 1-9
            for (int i = 1; i <= 9; i++)
            {
                sheet1.CreateRow(i).CreateCell(0).SetCellValue(i);
            }
            //create the cell formula
            for (int iRow = 1; iRow <= 9; iRow++)
            {
                IRow row = sheet1.GetRow(iRow);
                for (int iCol = 1; iCol <= 9; iCol++)
                {
                    //the first cell of each row * the first cell of each column
                    string formula = GetCellPosition(iRow, 0) + "*" + GetCellPosition(0, iCol);
                    row.CreateCell(iCol).CellFormula = formula;
                }
            }

            downLoadFile_HSSF(workBook, "Excel_2003_Plication.xls");
        }

        private string GetCellPosition(int row, int col)
        {
            col = Convert.ToInt32('A') + col;
            row = row + 1;
            return ((char)col) + row.ToString();
        }
        #endregion

        //带有数字格式的Excel
        protected void lnkbtnExcel_2003_NumberFormat_Click(object sender, EventArgs e)
        {
            HSSFWorkbook workBook = InitializeWorkbook_HSSF();

            ISheet sheet = workBook.CreateSheet("new sheet");
            //increase the width of Column A
            sheet.SetColumnWidth(0, 5000);
            //create the format instance
            IDataFormat format = workBook.CreateDataFormat();

            // Create a row and put some cells in it. Rows are 0 based.
            ICell cell = sheet.CreateRow(0).CreateCell(0);
            //set value for the cell
            cell.SetCellValue(1.2);
            //number format with 2 digits after the decimal point - "1.20"
            ICellStyle cellStyle = workBook.CreateCellStyle();
            cellStyle.DataFormat = HSSFDataFormat.GetBuiltinFormat("0.00");
            cell.CellStyle = cellStyle;

            //RMB currency format with comma    -   "¥20,000"
            ICell cell2 = sheet.CreateRow(1).CreateCell(0);
            cell2.SetCellValue(20000);
            ICellStyle cellStyle2 = workBook.CreateCellStyle();
            cellStyle2.DataFormat = format.GetFormat("¥#,##0");
            cell2.CellStyle = cellStyle2;

            //scentific number format   -   "3.15E+00"
            ICell cell3 = sheet.CreateRow(2).CreateCell(0);
            cell3.SetCellValue(3.151234);
            ICellStyle cellStyle3 = workBook.CreateCellStyle();
            cellStyle3.DataFormat = HSSFDataFormat.GetBuiltinFormat("0.00E+00");
            cell3.CellStyle = cellStyle3;

            //percent format, 2 digits after the decimal point    -  "99.33%"
            ICell cell4 = sheet.CreateRow(3).CreateCell(0);
            cell4.SetCellValue(0.99333);
            ICellStyle cellStyle4 = workBook.CreateCellStyle();
            cellStyle4.DataFormat = HSSFDataFormat.GetBuiltinFormat("0.00%");
            cell4.CellStyle = cellStyle4;

            //phone number format - "021-65881234"
            ICell cell5 = sheet.CreateRow(4).CreateCell(0);
            cell5.SetCellValue(02165881234);
            ICellStyle cellStyle5 = workBook.CreateCellStyle();
            cellStyle5.DataFormat = format.GetFormat("000-00000000");
            cell5.CellStyle = cellStyle5;

            //Chinese capitalized character number - 壹贰叁 元
            ICell cell6 = sheet.CreateRow(5).CreateCell(0);
            cell6.SetCellValue(123);
            ICellStyle cellStyle6 = workBook.CreateCellStyle();
            cellStyle6.DataFormat = format.GetFormat("[DbNum2][$-804]0 元");
            cell6.CellStyle = cellStyle6;

            //Chinese date string
            ICell cell7 = sheet.CreateRow(6).CreateCell(0);
            cell7.SetCellValue(new DateTime(2004, 5, 6));
            ICellStyle cellStyle7 = workBook.CreateCellStyle();
            cellStyle7.DataFormat = format.GetFormat("yyyy年m月d日");
            cell7.CellStyle = cellStyle7;


            //Chinese date string
            ICell cell8 = sheet.CreateRow(7).CreateCell(0);
            cell8.SetCellValue(new DateTime(2005, 11, 6));
            ICellStyle cellStyle8 = workBook.CreateCellStyle();
            cellStyle8.DataFormat = format.GetFormat("yyyy年m月d日");
            cell8.CellStyle = cellStyle8;

            downLoadFile_HSSF(workBook, "Excel_2003_NumberFormat.xls");
        }


        //带有保护机制的Excel
        protected void lnkbtnExcel_2003_ProtectSheet_Click(object sender, EventArgs e)
        {
            HSSFWorkbook workBook = InitializeWorkbook_HSSF();

            HSSFSheet sheet1 = (HSSFSheet)workBook.CreateSheet("Sheet1");

            ICell cell1 = sheet1.CreateRow(0).CreateCell(0);
            cell1.SetCellValue("This is a Sample");
            ICellStyle cs1 = workBook.CreateCellStyle();
            cs1.IsLocked = true;
            cell1.CellStyle = cs1;

            sheet1.ProtectSheet("test");

            downLoadFile_HSSF(workBook, "Excel_2003_ProtectSheet.xls");
        }


        //带有重复行和列的Excel
        protected void lnkbtnExcel_2003_RepeatingRowsAndColumns_Click(object sender, EventArgs e)
        {
            HSSFWorkbook workBook = InitializeWorkbook_HSSF();

            ISheet sheet1 = workBook.CreateSheet("first sheet");
            workBook.CreateSheet("second sheet");
            workBook.CreateSheet("third sheet");

            IFont boldFont = workBook.CreateFont();
            boldFont.FontHeightInPoints = 22;
            boldFont.Boldweight = (short)FontBoldWeight.Bold;

            ICellStyle boldStyle = workBook.CreateCellStyle();
            boldStyle.SetFont(boldFont);

            IRow row = sheet1.CreateRow(1);
            ICell cell = row.CreateCell(0);
            cell.SetCellValue("This quick brown fox");
            cell.CellStyle = (boldStyle);

            // Set the columns to repeat from column 0 to 2 on the first sheet
            workBook.SetRepeatingRowsAndColumns(0, 0, 2, -1, -1);
            // Set the rows to repeat from row 0 to 2 on the second sheet.
            workBook.SetRepeatingRowsAndColumns(1, -1, -1, 0, 2);
            // Set the the repeating rows and columns on the third sheet.
            workBook.SetRepeatingRowsAndColumns(2, 4, 5, 1, 2);

            downLoadFile_HSSF(workBook, "Excel_2003_RepeatingRowsAndColumns.xls");
        }

        //带有文字转换的Excel
        protected void lnkbtnExcel_2003_RotateText_Click(object sender, EventArgs e)
        {
            HSSFWorkbook workBook = InitializeWorkbook_HSSF();
            ISheet sheet1 = workBook.CreateSheet("Sheet1");

            //a valid rotate value is from -90 - 90
            int x = -90;
            for (int i = 1; i <= 13; i++)
            {
                IRow row = sheet1.CreateRow(i);
                for (int j = 0; j < 13; j++)
                {
                    //set the value
                    row.CreateCell(j).SetCellValue(x);
                    //set the style
                    ICellStyle style = workBook.CreateCellStyle();
                    style.Rotation = (short)x;
                    row.GetCell(j).CellStyle = style;
                    //increase x
                    x++;
                }
            }
            downLoadFile_HSSF(workBook, "Excel_2003_RotateText.xls");
        }


        #region 带有设置单元格活动范围的Excel
        protected void lnkbtnExcel_2003_SetActiveCellRange_Click(object sender, EventArgs e)
        {
            HSSFWorkbook workBook = InitializeWorkbook_HSSF();

            //use HSSFCell.SetAsActiveCell() to select B6 as the active column
            ISheet sheet1 = workBook.CreateSheet("ISheet A");
            CreateCellArray(sheet1);
            sheet1.GetRow(5).GetCell(1).SetAsActiveCell();
            //set TopRow and LeftCol to make B6 the first cell in the visible area
            sheet1.TopRow = 5;
            sheet1.LeftCol = 1;

            //use ISheet.SetActiveCell(), the sheet can be empty
            ISheet sheet2 = workBook.CreateSheet("ISheet B");
            sheet2.SetActiveCell(1, 5);

            //use ISheet.SetActiveCellRange to select a cell range
            ISheet sheet3 = workBook.CreateSheet("ISheet C");
            CreateCellArray(sheet3);
            sheet3.SetActiveCellRange(2, 20, 1, 50);
            //set the ISheet C as the active sheet
            workBook.SetActiveSheet(2);

            //use ISheet.SetActiveCellRange to select multiple cell ranges
            ISheet sheet4 = workBook.CreateSheet("ISheet D");
            CreateCellArray(sheet4);
            List<CellRangeAddress8Bit> cellranges = new List<CellRangeAddress8Bit>();
            cellranges.Add(new CellRangeAddress8Bit(1, 5, 10, 100));
            cellranges.Add(new CellRangeAddress8Bit(6, 7, 8, 9));
            sheet4.SetActiveCellRange(cellranges, 1, 6, 9);

            downLoadFile_HSSF(workBook, "Excel_2003_SetActiveCellRange.xls");
        }

        static void CreateCellArray(ISheet sheet)
        {
            for (int i = 0; i < 300; i++)
            {
                IRow row = sheet.CreateRow(i);
                for (int j = 0; j < 150; j++)
                {
                    ICell cell = row.CreateCell(j);
                    cell.SetCellValue(i * j);
                }
            }
        }

        #endregion

        //设置对齐的Excel
        protected void lnkbtnExcel_2003_SetAlignment_Click(object sender, EventArgs e)
        {
            HSSFWorkbook workBook = InitializeWorkbook_HSSF();

            ISheet sheet1 = workBook.CreateSheet("Sheet1");

            //set the column width respectively
            sheet1.SetColumnWidth(0, 3000);
            sheet1.SetColumnWidth(1, 3000);
            sheet1.SetColumnWidth(2, 3000);

            for (int i = 1; i <= 10; i++)
            {
                //create the row
                IRow row = sheet1.CreateRow(i);
                //set height of the row
                row.HeightInPoints = 100;

                //create the first cell
                row.CreateCell(0).SetCellValue("Left");
                ICellStyle styleLeft = workBook.CreateCellStyle();
                styleLeft.Alignment = HorizontalAlignment.Left;
                styleLeft.VerticalAlignment = VerticalAlignment.Top;
                row.GetCell(0).CellStyle = styleLeft;
                //set indention for the text in the cell
                styleLeft.Indention = 3;

                //create the second cell
                row.CreateCell(1).SetCellValue("Center Hello World Hello WorldHello WorldHello WorldHello WorldHello World");
                ICellStyle styleMiddle = workBook.CreateCellStyle();
                styleMiddle.Alignment = HorizontalAlignment.Center;
                styleMiddle.VerticalAlignment = VerticalAlignment.Center;
                row.GetCell(1).CellStyle = styleMiddle;
                //wrap the text in the cell
                styleMiddle.WrapText = true;


                //create the third cell
                row.CreateCell(2).SetCellValue("Right");
                ICellStyle styleRight = workBook.CreateCellStyle();
                styleRight.Alignment = HorizontalAlignment.Justify;
                styleRight.VerticalAlignment = VerticalAlignment.Bottom;
                row.GetCell(2).CellStyle = styleRight;

            }

            downLoadFile_HSSF(workBook, "Excel_2003_SetAlignment.xls");
        }

        //设置边框区域的Excel
        protected void lnkbtnExcel_2003_SetBordersOfRegion_Click(object sender, EventArgs e)
        {
            HSSFWorkbook workBook = InitializeWorkbook_HSSF();
            ISheet sheet1 = workBook.CreateSheet("Sheet1");
            //create a common style
            ICellStyle blackBorder = workBook.CreateCellStyle();
            blackBorder.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            blackBorder.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            blackBorder.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            blackBorder.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            blackBorder.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            blackBorder.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            blackBorder.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            blackBorder.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;

            //create horizontal 1-9
            for (int i = 1; i <= 9; i++)
            {
                sheet1.CreateRow(0).CreateCell(i).SetCellValue(i);
            }
            //create vertical 1-9
            for (int i = 1; i <= 9; i++)
            {
                sheet1.CreateRow(i).CreateCell(0).SetCellValue(i);
            }
            //create the cell formula
            for (int iRow = 1; iRow <= 9; iRow++)
            {
                IRow row = sheet1.GetRow(iRow);
                for (int iCol = 1; iCol <= 9; iCol++)
                {
                    //the first cell of each row * the first cell of each column
                    string formula = GetCellPosition(iRow, 0) + "*" + GetCellPosition(0, iCol);
                    ICell cell = row.CreateCell(iCol);
                    cell.CellFormula = formula;
                    //set the cellstyle to the cell
                    cell.CellStyle = blackBorder;
                }
            }

            downLoadFile_HSSF(workBook, "Excel_2003_SetBordersOfRegion.xls");
        }

        //设置边框样式的Excel
        protected void lnkbtnExcel_2003_SetBorderStyle_Click(object sender, EventArgs e)
        {
            HSSFWorkbook workBook = InitializeWorkbook_HSSF();
            ISheet sheet = workBook.CreateSheet("new sheet");

            // Create a row and put some cells in it. Rows are 0 based.
            IRow row = sheet.CreateRow(1);

            // Create a cell and put a value in it.
            ICell cell = row.CreateCell(1);
            cell.SetCellValue(4);

            // Style the cell with borders all around.
            ICellStyle style = workBook.CreateCellStyle();
            style.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            style.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            style.BorderLeft = NPOI.SS.UserModel.BorderStyle.DashDotDot;
            style.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Green.Index;
            style.BorderRight = NPOI.SS.UserModel.BorderStyle.Hair;
            style.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Blue.Index;
            style.BorderTop = NPOI.SS.UserModel.BorderStyle.MediumDashed;
            style.TopBorderColor = HSSFColor.Orange.Index;

            style.BorderDiagonal = NPOI.SS.UserModel.BorderDiagonal.Forward;
            style.BorderDiagonalColor = NPOI.HSSF.Util.HSSFColor.Gold.Index;
            style.BorderDiagonalLineStyle = NPOI.SS.UserModel.BorderStyle.Medium;

            cell.CellStyle = style;
            // Create a cell and put a value in it.
            ICell cell2 = row.CreateCell(2);
            cell2.SetCellValue(5);
            ICellStyle style2 = workBook.CreateCellStyle();
            style2.BorderDiagonal = BorderDiagonal.Backward;
            style2.BorderDiagonalColor = NPOI.HSSF.Util.HSSFColor.Red.Index;
            style2.BorderDiagonalLineStyle = NPOI.SS.UserModel.BorderStyle.Medium;
            cell2.CellStyle = style2;

            downLoadFile_HSSF(workBook, "Excel_2003_SetBorderStyle.xls");
        }

        //设置单元格的评论/批注的Excel
        protected void lnkbtnExcel_2003_SetCellComment_Click(object sender, EventArgs e)
        {
            HSSFWorkbook workBook = InitializeWorkbook_HSSF();

            ISheet sheet = workBook.CreateSheet("ICell comments in POI HSSF");

            // Create the drawing patriarch. This is the top level container for all shapes including cell comments.
            IDrawing patr = (HSSFPatriarch)sheet.CreateDrawingPatriarch();

            //Create a cell in row 3
            ICell cell1 = sheet.CreateRow(3).CreateCell(1);
            cell1.SetCellValue(new HSSFRichTextString("Hello, World"));

            //anchor defines size and position of the comment in worksheet
            IComment comment1 = patr.CreateCellComment(new HSSFClientAnchor(0, 0, 0, 0, 4, 2, 6, 5));

            // set text in the comment
            comment1.String = (new HSSFRichTextString("We can set comments in POI"));

            //set comment author.
            //you can see it in the status bar when moving mouse over the commented cell
            comment1.Author = ("Apache Software Foundation");

            // The first way to assign comment to a cell is via HSSFCell.SetCellComment method
            cell1.CellComment = (comment1);

            //Create another cell in row 6
            ICell cell2 = sheet.CreateRow(6).CreateCell(1);
            cell2.SetCellValue(36.6);


            HSSFComment comment2 = (HSSFComment)patr.CreateCellComment(new HSSFClientAnchor(0, 0, 0, 0, 4, 8, 6, 11));
            //modify background color of the comment
            comment2.SetFillColor(204, 236, 255);

            HSSFRichTextString str = new HSSFRichTextString("Normal body temperature");

            //apply custom font to the text in the comment
            IFont font = workBook.CreateFont();
            font.FontName = ("Arial");
            font.FontHeightInPoints = 10;
            font.Boldweight = (short)FontBoldWeight.Bold;
            font.Color = HSSFColor.Red.Index;
            str.ApplyFont(font);

            comment2.String = str;
            comment2.Visible = true; //by default comments are hidden. This one is always visible.

            comment2.Author = "Bill Gates";

            /**
             * The second way to assign comment to a cell is to implicitly specify its row and column.
             * Note, it is possible to set row and column of a non-existing cell.
             * It works, the commnet is visible.
             */
            comment2.Row = 6;
            comment2.Column = 1;

            downLoadFile_HSSF(workBook, "Excel_2003_SetBorderStyle.xls");
        }

        //设置单元格的值的Excel
        protected void lnkbtnExcel_2003_SetCellValues_Click(object sender, EventArgs e)
        {
            HSSFWorkbook workBook = InitializeWorkbook_HSSF();
            ISheet sheet1 = workBook.CreateSheet("Sheet1");

            sheet1.CreateRow(0).CreateCell(0).SetCellValue("This is a Sample");
            int x = 1;
            for (int i = 1; i <= 15; i++)
            {
                IRow row = sheet1.CreateRow(i);
                for (int j = 0; j < 15; j++)
                {
                    row.CreateCell(j).SetCellValue(x++);
                }
            }

            downLoadFile_HSSF(workBook, "Excel_2003_SetCellValues.xls");
        }

        //设置日期格式的单元格的Excel
        protected void lnkbtnExcel_2003_SetDateCell_Click(object sender, EventArgs e)
        {
            HSSFWorkbook workBook = InitializeWorkbook_HSSF();
            ISheet sheet = workBook.CreateSheet("new sheet");
            // Create a row and put some cells in it. Rows are 0 based.
            IRow row = sheet.CreateRow(0);

            // Create a cell and put a date value in it.  The first cell is not styled as a date.
            ICell cell = row.CreateCell(0);
            cell.SetCellValue(DateTime.Now);

            // we style the second cell as a date (and time).  It is important to Create a new cell style from the workbook
            // otherwise you can end up modifying the built in style and effecting not only this cell but other cells.
            ICellStyle cellStyle = workBook.CreateCellStyle();

            // Perhaps this may only works for Chinese date, I don't have english office on hand
            cellStyle.DataFormat = workBook.CreateDataFormat().GetFormat("[$-409]h:mm:ss AM/PM;@");
            cell.CellStyle = cellStyle;

            //set chinese date format
            ICell cell2 = row.CreateCell(1);
            cell2.SetCellValue(new DateTime(2008, 5, 5));
            ICellStyle cellStyle2 = workBook.CreateCellStyle();
            IDataFormat format = workBook.CreateDataFormat();
            cellStyle2.DataFormat = format.GetFormat("yyyy年m月d日");
            cell2.CellStyle = cellStyle2;

            ICell cell3 = row.CreateCell(2);
            cell3.CellFormula = "DateValue(\"2005-11-11 11:11:11\")";
            ICellStyle cellStyle3 = workBook.CreateCellStyle();
            cellStyle3.DataFormat = HSSFDataFormat.GetBuiltinFormat("m/d/yy h:mm");
            cell3.CellStyle = cellStyle3;

            downLoadFile_HSSF(workBook, "Excel_2003_SetDateCell.xls");
        }

        //设置打印区域的Excel
        protected void lnkbtnExcel_2003_SetPrintArea_Click(object sender, EventArgs e)
        {
            HSSFWorkbook workBook = InitializeWorkbook_HSSF();

            ISheet sheet1 = workBook.CreateSheet("Sheet1");

            //fill some data
            sheet1.CreateRow(0).CreateCell(0).SetCellValue("This is a Sample");
            int x = 1;
            int i = 1;
            for (i = 1; i <= 15; i++)
            {
                IRow row = sheet1.CreateRow(i);
                for (int j = 0; j < 15; j++)
                {
                    row.CreateCell(j).SetCellValue(x++);
                }
            }

            //set print area
            workBook.SetPrintArea(0, "A5:G20");

            downLoadFile_HSSF(workBook, "Excel_2003_SetPrintArea.xls");
        }

        //设置打印设置的Excel
        protected void lnkbtnExcel_2003_SetPrintSettings_Click(object sender, EventArgs e)
        {
            HSSFWorkbook workBook = InitializeWorkbook_HSSF();

            ISheet sheet1 = workBook.CreateSheet("Sheet1");
            sheet1.SetMargin(MarginType.RightMargin, (double)0.5);
            sheet1.SetMargin(MarginType.TopMargin, (double)0.6);
            sheet1.SetMargin(MarginType.LeftMargin, (double)0.4);
            sheet1.SetMargin(MarginType.BottomMargin, (double)0.3);


            sheet1.PrintSetup.Copies = 3;
            sheet1.PrintSetup.NoColor = true;
            sheet1.PrintSetup.Landscape = true;
            sheet1.PrintSetup.PaperSize = (short)PaperSize.A4;

            sheet1.FitToPage = true;
            sheet1.PrintSetup.FitHeight = 2;
            sheet1.PrintSetup.FitWidth = 3;
            sheet1.IsPrintGridlines = true;

            sheet1.CreateRow(0).CreateCell(0).SetCellValue("This is a Sample");
            int x = 1;
            for (int i = 1; i <= 15; i++)
            {
                IRow row = sheet1.CreateRow(i);
                for (int j = 0; j < 15; j++)
                {
                    row.CreateCell(j).SetCellValue(x++);
                }
            }

            ISheet sheet2 = workBook.CreateSheet("Sheet2");
            sheet2.PrintSetup.Copies = 1;
            sheet2.PrintSetup.Landscape = false;
            sheet2.PrintSetup.Notes = true;
            sheet2.PrintSetup.EndNote = true;
            sheet2.PrintSetup.CellError = DisplayCellErrorType.ErrorAsNA;
            sheet2.PrintSetup.PaperSize = (short)PaperSize.A5;

            x = 100;
            for (int i = 1; i <= 15; i++)
            {
                IRow row = sheet2.CreateRow(i);
                for (int j = 0; j < 15; j++)
                {
                    row.CreateCell(j).SetCellValue(x++);
                }
            }

            downLoadFile_HSSF(workBook, "Excel_2003_SetPrintSettings.xls");
        }

        //设置宽度和高度的Excel
        protected void lnkbtnExcel_2003_SetWidthAndHeight_Click(object sender, EventArgs e)
        {
            HSSFWorkbook workBook = InitializeWorkbook_HSSF();
            ISheet sheet1 = workBook.CreateSheet("Sheet1");
            //set the width of columns
            sheet1.SetColumnWidth(0, 50 * 256);
            sheet1.SetColumnWidth(1, 100 * 256);
            sheet1.SetColumnWidth(2, 150 * 256);

            //set the width of height
            sheet1.CreateRow(0).Height = 100 * 20;
            sheet1.CreateRow(1).Height = 200 * 20;
            sheet1.CreateRow(2).Height = 300 * 20;

            sheet1.DefaultRowHeightInPoints = 50;
            downLoadFile_HSSF(workBook, "Excel_2003_SetWidthAndHeight.xls");
        }

        //缩小到合适列的Excel
        protected void lnkbtnExcel_2003_ShrinkToFitColumn_Click(object sender, EventArgs e)
        {
            HSSFWorkbook workBook = InitializeWorkbook_HSSF();

            ISheet sheet = workBook.CreateSheet("Sheet1");
            IRow row = sheet.CreateRow(0);
            //create cell value
            ICell cell1 = row.CreateCell(0);
            cell1.SetCellValue("This is a test");
            //apply ShrinkToFit to cellstyle
            ICellStyle cellstyle1 = workBook.CreateCellStyle();
            cellstyle1.ShrinkToFit = true;
            cell1.CellStyle = cellstyle1;
            //create cell value
            row.CreateCell(1).SetCellValue("Hello World");
            row.GetCell(1).CellStyle = cellstyle1;

            downLoadFile_HSSF(workBook, "Excel_2003_ShrinkToFitColumn.xls");
        }

        //带有拆分和冻结窗格的Excel
        protected void lnkbtnExcel_2003_SplitAndFreezePanes_Click(object sender, EventArgs e)
        {
            HSSFWorkbook workBook = InitializeWorkbook_HSSF();

            ISheet sheet1 = workBook.CreateSheet("new sheet");
            ISheet sheet2 = workBook.CreateSheet("second sheet");
            ISheet sheet3 = workBook.CreateSheet("third sheet");
            ISheet sheet4 = workBook.CreateSheet("fourth sheet");

            // Freeze just one row
            sheet1.CreateFreezePane(0, 1, 0, 1);
            // Freeze just one column
            sheet2.CreateFreezePane(1, 0, 1, 0);
            // Freeze the columns and rows (forget about scrolling position of the lower right quadrant).
            sheet3.CreateFreezePane(2, 2);
            // Create a split with the lower left side being the active quadrant
            sheet4.CreateSplitPane(2000, 2000, 0, 0, PanePosition.LowerLeft);

            downLoadFile_HSSF(workBook, "Excel_2003_SplitAndFreezePanes.xls");
        }

        #region 带有时间片演示的Excel
        protected void lnkbtnExcel_2003_TimeSheetDemo_Click(object sender, EventArgs e)
        {
            HSSFWorkbook workBook = InitializeWorkbook_HSSF();

            Dictionary<String, ICellStyle> styles = CreateStyles_timeDemo(workBook);

            String[] titles = {
            "Person",	"ID", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun",
            "Total\nHrs", "Overtime\nHrs", "Regular\nHrs"
        };
            Object[,] sample_data = {
            {"Yegor Kozlov", "YK", 5.0, 8.0, 10.0, 5.0, 5.0, 7.0, 6.0},
            {"Gisella Bronzetti", "GB", 4.0, 3.0, 1.0, 3.5, null, null, 4.0}
        };

            ISheet sheet = workBook.CreateSheet("Timesheet");
            IPrintSetup printSetup = sheet.PrintSetup;
            printSetup.Landscape = true;
            sheet.FitToPage = (true);
            sheet.HorizontallyCenter = (true);

            //title row
            IRow titleRow = sheet.CreateRow(0);
            titleRow.HeightInPoints = (45);
            ICell titleCell = titleRow.CreateCell(0);
            titleCell.SetCellValue("Weekly Timesheet");
            titleCell.CellStyle = (styles["title"]);
            sheet.AddMergedRegion(CellRangeAddress.ValueOf("$A$1:$L$1"));

            //header row
            IRow headerRow = sheet.CreateRow(1);
            headerRow.HeightInPoints = (40);
            ICell headerCell;
            for (int i = 0; i < titles.Length; i++)
            {
                headerCell = headerRow.CreateCell(i);
                headerCell.SetCellValue(titles[i]);
                headerCell.CellStyle = (styles["header"]);
            }


            int rownum = 2;
            for (int i = 0; i < 10; i++)
            {
                IRow row = sheet.CreateRow(rownum++);
                for (int j = 0; j < titles.Length; j++)
                {
                    ICell cell = row.CreateCell(j);
                    if (j == 9)
                    {
                        //the 10th cell contains sum over week days, e.g. SUM(C3:I3)
                        String reference = "C" + rownum + ":I" + rownum;
                        cell.CellFormula = ("SUM(" + reference + ")");
                        cell.CellStyle = (styles["formula"]);
                    }
                    else if (j == 11)
                    {
                        cell.CellFormula = ("J" + rownum + "-K" + rownum);
                        cell.CellStyle = (styles["formula"]);
                    }
                    else
                    {
                        cell.CellStyle = (styles["cell"]);
                    }
                }
            }

            //row with totals below
            IRow sumRow = sheet.CreateRow(rownum++);
            sumRow.HeightInPoints = (35);
            ICell cell1 = sumRow.CreateCell(0);
            cell1.CellStyle = (styles["formula"]);

            ICell cell2 = sumRow.CreateCell(1);
            cell2.SetCellValue("Total Hrs:");
            cell2.CellStyle = (styles["formula"]);

            for (int j = 2; j < 12; j++)
            {
                ICell cell = sumRow.CreateCell(j);
                String reference = (char)('A' + j) + "3:" + (char)('A' + j) + "12";
                cell.CellFormula = ("SUM(" + reference + ")");
                if (j >= 9)
                    cell.CellStyle = (styles["formula_2"]);
                else
                    cell.CellStyle = (styles["formula"]);
            }

            rownum++;
            sumRow = sheet.CreateRow(rownum++);
            sumRow.HeightInPoints = 25;
            ICell cell3 = sumRow.CreateCell(0);
            cell3.SetCellValue("Total Regular Hours");
            cell3.CellStyle = styles["formula"];
            cell3 = sumRow.CreateCell(1);
            cell3.CellFormula = ("L13");
            cell3.CellStyle = styles["formula_2"];
            sumRow = sheet.CreateRow(rownum++);
            sumRow.HeightInPoints = (25);
            cell3 = sumRow.CreateCell(0);
            cell3.SetCellValue("Total Overtime Hours");
            cell3.CellStyle = styles["formula"];
            cell3 = sumRow.CreateCell(1);
            cell3.CellFormula = ("K13");
            cell3.CellStyle = styles["formula_2"];

            //set sample data
            for (int i = 0; i < sample_data.GetLength(0); i++)
            {
                IRow row = sheet.GetRow(2 + i);
                for (int j = 0; j < sample_data.GetLength(1); j++)
                {
                    if (sample_data[i, j] == null)
                        continue;

                    if (sample_data[i, j] is String)
                    {
                        row.GetCell(j).SetCellValue((String)sample_data[i, j]);
                    }
                    else
                    {
                        row.GetCell(j).SetCellValue((Double)sample_data[i, j]);
                    }
                }
            }

            //finally set column widths, the width is measured in units of 1/256th of a character width
            sheet.SetColumnWidth(0, 30 * 256); //30 characters wide
            for (int i = 2; i < 9; i++)
            {
                sheet.SetColumnWidth(i, 6 * 256);  //6 characters wide
            }
            sheet.SetColumnWidth(10, 10 * 256); //10 characters wide

            downLoadFile_HSSF(workBook, "Excel_2003_TimeSheetDemo.xls");
        }

        private Dictionary<String, ICellStyle> CreateStyles_timeDemo(IWorkbook wb)
        {
            Dictionary<String, ICellStyle> styles = new Dictionary<String, ICellStyle>();
            ICellStyle style;
            IFont titleFont = wb.CreateFont();
            titleFont.FontHeightInPoints = ((short)18);
            titleFont.Boldweight = (short)FontBoldWeight.Bold;
            style = wb.CreateCellStyle();
            style.Alignment = HorizontalAlignment.Center;
            style.VerticalAlignment = VerticalAlignment.Center;
            style.SetFont(titleFont);
            styles.Add("title", style);

            IFont monthFont = wb.CreateFont();
            monthFont.FontHeightInPoints = ((short)11);
            monthFont.Color = (IndexedColors.White.Index);
            style = wb.CreateCellStyle();
            style.Alignment = HorizontalAlignment.Center;
            style.VerticalAlignment = VerticalAlignment.Center;
            style.FillForegroundColor = (IndexedColors.Grey50Percent.Index);
            style.FillPattern = FillPattern.SolidForeground;
            style.SetFont(monthFont);
            style.WrapText = (true);
            styles.Add("header", style);

            style = wb.CreateCellStyle();
            style.Alignment = HorizontalAlignment.Center;
            style.WrapText = (true);
            style.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            style.RightBorderColor = (IndexedColors.Black.Index);
            style.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            style.LeftBorderColor = (IndexedColors.Black.Index);
            style.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            style.TopBorderColor = (IndexedColors.Black.Index);
            style.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            style.BottomBorderColor = (IndexedColors.Black.Index);
            styles.Add("cell", style);

            style = wb.CreateCellStyle();
            style.Alignment = HorizontalAlignment.Center;
            style.VerticalAlignment = VerticalAlignment.Center;
            style.FillForegroundColor = (IndexedColors.Grey25Percent.Index);
            style.FillPattern = FillPattern.SolidForeground;
            style.DataFormat = (wb.CreateDataFormat().GetFormat("0.00"));
            styles.Add("formula", style);

            style = wb.CreateCellStyle();
            style.Alignment = HorizontalAlignment.Center;
            style.VerticalAlignment = VerticalAlignment.Center;
            style.FillForegroundColor = (IndexedColors.Grey40Percent.Index);
            style.FillPattern = FillPattern.SolidForeground;
            style.DataFormat = wb.CreateDataFormat().GetFormat("0.00");
            styles.Add("formula_2", style);

            return styles;
        }

        #endregion


        //带有基本公式的Excel
        protected void lnkbtnExcel_2003_UseBasicFormula_Click(object sender, EventArgs e)
        {
            HSSFWorkbook workBook = InitializeWorkbook_HSSF();

            ISheet s1 = workBook.CreateSheet("Sheet1");
            //set A2
            s1.CreateRow(1).CreateCell(0).SetCellValue(-5);
            //set B2
            s1.GetRow(1).CreateCell(1).SetCellValue(1111);
            //set C2
            s1.GetRow(1).CreateCell(2).SetCellValue(7.623);
            //set A3
            s1.CreateRow(2).CreateCell(0).SetCellValue(2.2);

            //set A4=A2+A3
            s1.CreateRow(3).CreateCell(0).CellFormula = "A2+A3";
            //set D2=SUM(A2:C2);
            s1.GetRow(1).CreateCell(3).CellFormula = "SUM(A2:C2)";
            //set A5=cos(5)+sin(10)
            s1.CreateRow(4).CreateCell(0).CellFormula = "cos(5)+sin(10)";


            //create another sheet
            ISheet s2 = workBook.CreateSheet("Sheet2");
            //set cross-sheet reference
            s2.CreateRow(0).CreateCell(0).CellFormula = "Sheet1!A2+Sheet1!A3";

            downLoadFile_HSSF(workBook, "Excel_2003_UseBasicFormula.xls");
        }

        //在单元格中使用换行的Excel
        protected void lnkbtnExcel_2003_UseNewLinesInCells_Click(object sender, EventArgs e)
        {
            HSSFWorkbook workBook = InitializeWorkbook_HSSF();

            ISheet sheet1 = workBook.CreateSheet("Sheet1");

            //use newlines in cell
            IRow row1 = sheet1.CreateRow(0);
            ICell cell1 = row1.CreateCell(0);

            //to enable newlines you need set a cell styles with wrap=true
            ICellStyle cs = workBook.CreateCellStyle();
            cs.WrapText = true;
            cell1.CellStyle = cs;

            //increase row height to accomodate two lines of text
            row1.HeightInPoints = 3 * sheet1.DefaultRowHeightInPoints;
            cell1.SetCellValue("This is a \n Hello \n World!");

            downLoadFile_HSSF(workBook, "Excel_2003_UseNewLinesInCells.xls");
        }

        //带有放大缩小的Sheet的Excel
        protected void lnkbtnExcel_2003_ZoomSheet_Click(object sender, EventArgs e)
        {
            HSSFWorkbook workBook = InitializeWorkbook_HSSF();

            ISheet sheet1 = workBook.CreateSheet("new sheet");
            sheet1.SetZoom(3, 4);   // 75 percent magnification

            downLoadFile_HSSF(workBook, "Excel_2003_ZoomSheet.xls");
        }

        //带有可选择性设置密码的Excel
        protected void lnkbtnExcel_2003_SetPassword_Click(object sender, EventArgs e)
        {
            //1.创建工作簿
            HSSFWorkbook workBook = new HSSFWorkbook();//会自动生成sheet
            //IWorkbook workBook = new HSSFWorkbook();//不会生成sheet

            //2..创建sheet2
            //HSSFSheet workSheet = (NPOI.HSSF.UserModel.HSSFSheet)workBook.CreateSheet("sheet1");
            //下面也是创建sheet的步骤
            ISheet workSheet = workBook.CreateSheet("sheet1");
            #region

            //注意Excel的sheet中的行数其实是有相应的最大值的是 65535行，当超出这个范围的时候就会报错 应添加一个新的sheet
            for (int i = 0; i < 1000; i++)
            {
                if (i == 0)
                {
                    //设置表头

                    //3.创建行
                    IRow row = workSheet.CreateRow(i);

                    //设置颜色和style
                    row.Height = 30 * 20;//这里的30才是真正长度
                    workSheet.SetColumnWidth(0, 60 * 256);//单元格的下标,宽度60才是真正的宽度
                    ICellStyle style = workBook.CreateCellStyle();

                    IFont font = workBook.CreateFont();
                    font.FontName = "微软雅黑";
                    font.FontHeightInPoints = 10;
                    font.Color = NPOI.HSSF.Util.HSSFColor.Blue.Index;
                    font.IsItalic = true;//下划线  
                    style.SetFont(font);
                    style.FillBackgroundColor = 53;
                    style.FillForegroundColor = 53;
                    style.FillPattern = FillPattern.NoFill;
                    style.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Yellow.Index;
                    //style.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
                    style.FillPattern = FillPattern.SolidForeground; //前景色填充


                    //4. 创建单元格
                    ICell cell_date = row.CreateCell(0);
                    ICell cell_Line = row.CreateCell(1);
                    ICell cell_Finder = row.CreateCell(2);
                    ICell cell_Area = row.CreateCell(3);
                    ICell cell_QueXianMiaoShu = row.CreateCell(4);
                    ICell cell_YiChangLeiXing = row.CreateCell(5);
                    ICell cell_IsCil = row.CreateCell(6);
                    ICell cell_Action = row.CreateCell(7);
                    ICell cell_SheBeiFuZeRen = row.CreateCell(8);
                    ICell cell_IsRepeat = row.CreateCell(9);
                    ICell cell_Help = row.CreateCell(10);
                    ICell cell_PlanTime = row.CreateCell(11);
                    ICell cell_ActulyTime = row.CreateCell(12);
                    ICell cell_WanChengRen = row.CreateCell(13);
                    ICell cell_Status = row.CreateCell(14);
                    ICell cell_GengXinCIL = row.CreateCell(15);
                    ICell cell_Remark = row.CreateCell(16);


                    //5.赋值
                    cell_date.SetCellValue("日期");
                    cell_Line.SetCellValue("Line");
                    cell_Finder.SetCellValue("发现人");
                    cell_Area.SetCellValue("设备区域");
                    cell_QueXianMiaoShu.SetCellValue("缺陷描述");
                    cell_YiChangLeiXing.SetCellValue("异常类型");
                    cell_IsCil.SetCellValue("是否为CIL是发现");
                    cell_Action.SetCellValue("行动计划");
                    cell_SheBeiFuZeRen.SetCellValue("设备负责人");
                    cell_IsRepeat.SetCellValue("重复的/关键的");
                    cell_Help.SetCellValue("寻求资源帮助");
                    cell_PlanTime.SetCellValue("预计完成日");
                    cell_ActulyTime.SetCellValue("实际完成日");
                    cell_WanChengRen.SetCellValue("完成人");
                    cell_Status.SetCellValue("状态");
                    cell_GengXinCIL.SetCellValue("是否需要更新CIL");
                    cell_Remark.SetCellValue("备注");

                    //注意上锁必须卸载赋值后面 否则都为上锁
                    //猜想：假如不谢锁定，直接来一个ProtectSheet 会不会直接锁住  -----答案 会锁住
                    //cellStyle赋值 前面的会被后面的替换！

                    foreach (ICell cell in row)
                    {
                        cell.CellStyle = style;
                    }

                    //没上锁
                    ICellStyle unLocked = workBook.CreateCellStyle();
                    unLocked.IsLocked = false;
                    cell_GengXinCIL.CellStyle = unLocked;

                    //上锁
                    //ICellStyle locked = workBook.CreateCellStyle();
                    //locked.IsLocked = true;
                    ////workSheet.ProtectSheet("123456");
                    //cell_Remark.CellStyle = locked;
                    workSheet.ProtectSheet("123456");//设置密码

                    //workSheet.CreateFreezePane(1,1);//列,行 从第几行和第几列开始冻结  表示固定第一行
                    workSheet.CreateFreezePane(1, 1, 1, 1);//这个是 表示固定第一行和最左边的一列
                }
                else
                {


                    //设置表头

                    //3.创建行
                    IRow row = workSheet.CreateRow(i);

                    //4. 创建单元格
                    ICell cell_date = row.CreateCell(0);
                    ICell cell_Line = row.CreateCell(1);
                    ICell cell_Finder = row.CreateCell(2);
                    ICell cell_Area = row.CreateCell(3);
                    ICell cell_QueXianMiaoShu = row.CreateCell(4);
                    ICell cell_YiChangLeiXing = row.CreateCell(5);
                    ICell cell_IsCil = row.CreateCell(6);
                    ICell cell_Action = row.CreateCell(7);
                    ICell cell_SheBeiFuZeRen = row.CreateCell(8);
                    ICell cell_IsRepeat = row.CreateCell(9);
                    ICell cell_Help = row.CreateCell(10);
                    ICell cell_PlanTime = row.CreateCell(11);
                    ICell cell_ActulyTime = row.CreateCell(12);
                    ICell cell_WanChengRen = row.CreateCell(13);
                    ICell cell_Status = row.CreateCell(14);
                    ICell cell_GengXinCIL = row.CreateCell(15);
                    ICell cell_Remark = row.CreateCell(16);

                    //5.赋值
                    cell_date.SetCellValue(i);
                    cell_Line.SetCellValue(i + 1);
                    cell_Finder.SetCellValue(i + 2);
                    cell_Area.SetCellValue(i + 3);
                    cell_QueXianMiaoShu.SetCellValue(i + 4);
                    cell_YiChangLeiXing.SetCellValue(i + 5);
                    cell_IsCil.SetCellValue(i + 6);
                    cell_Action.SetCellValue(i + 7);
                    cell_SheBeiFuZeRen.SetCellValue(i + 8);
                    cell_IsRepeat.SetCellValue(i + 9);
                    cell_Help.SetCellValue(i + 10);
                    cell_PlanTime.SetCellValue(i + 11);
                    cell_ActulyTime.SetCellValue(i + 12);
                    cell_WanChengRen.SetCellValue(i + 13);
                    cell_Status.SetCellValue(i + 14);
                    cell_GengXinCIL.SetCellValue(i + 15);
                    cell_Remark.SetCellValue(i + 16);
                }
            }

            #endregion

            workSheet.PrintSetup.Landscape = true;//表示打印的时候 横向打印  FALSE为 纵向打印
            //还可以做其他的设置 打印设置 （缩放，纸张，页面宽高，网格线，单色打印，草稿品质，打印顺序，批注）等等  http://tonyqus.sinaapp.com/archives/271




            //6.保存  打开以一个新文件或者是创建一个新的文件写入  
            //注意必须要用using 或者是如下*********************************
            //using (FileStream file = File.OpenWrite(@"D:\excel.xls"))
            //{
            //    //7.向打开的这个xls文件中写入并保存。
            //    workBook.Write(file);

            //    Response.Write("<script>alert('导出Excel成功！')</script>");
            //}

            //FileStream file = File.OpenWrite(@"D:\excel1.xls");

            //7.向打开的这个xls文件中写入并保存。
            //workBook.Write(file);
            //file.Close();//注意别忘记关闭文件流
            //Response.Write("<script>alert('导出Excel成功！')</script>");


            //8.提供客户端下载

            //设置响应的类型为Excel
            Response.ContentType = "application/vnd.ms-excel";
            //设置下载的Excel文件名
            Response.AddHeader("Content-Disposition", string.Format("attachment;filename={0}", "Excel_2003_SetPassword.xls"));//目前只能为03的Excel
            //Clear方法删除所有缓存中的HTML输出。但此方法只删除Response显示输入信息，不删除Response头信息。以免影响导出数据的完整性。
            Response.Clear();

            //注意类型
            using (MemoryStream file = new MemoryStream())
            {
                //将工作簿的内容放到内存流中
                workBook.Write(file);
                //将内存流转换成字节数组发送到客户端
                /*
                 * //这种写法和下面的相比 会产生大量的个byte[]临时变量数据
                Response.BinaryWrite(file.GetBuffer());
                Response.End();
                */
                file.WriteTo(Response.OutputStream);
            }
        }

        #endregion

        #region OOXML  压缩文件
        protected void lnkbtn_CreateFile_Click(object sender, EventArgs e)
        {
            //create ooxml file in memory
            OPCPackage p = OPCPackage.Create(new MemoryStream());

            //create package parts
            PackagePartName pn1 = new PackagePartName(new Uri("/a/abcd/e", UriKind.Relative), true);
            if (!p.ContainPart(pn1))
                p.CreatePart(pn1, MediaTypeNames.Text.Plain);//文件里面内容为纯文本

            PackagePartName pn2 = new PackagePartName(new Uri("/b/test.xml", UriKind.Relative), true);
            if (!p.ContainPart(pn2))
                p.CreatePart(pn2, MediaTypeNames.Text.Xml);//文件里面的内容为xml

            //save file 
            p.Save("D:\\test_Create.zip");

            //don't forget to close it
            p.Close();
        }

        protected void lnkbtn_ModifyFile_Click(object sender, EventArgs e)
        {
            string path = Server.MapPath("ZipMode/test.zip");

            OPCPackage p = OPCPackage.Open(path, PackageAccess.READ_WRITE);

            PackagePartName pn3 = new PackagePartName(new Uri("/c.xml", UriKind.Relative), true);
            if (!p.ContainPart(pn3))
                p.CreatePart(pn3, MediaTypeNames.Text.Xml);

            //save file 
            p.Save("D:\\test_Midify.zip");

            //don't forget to close it
            p.Close();
        }

        #endregion

        #region Office属性设置  右键属性-->自定义

        //创建自定义属性
        protected void lnkbtn_CreateCustomProperties_Click(object sender, EventArgs e)
        {
            POIFSFileSystem fs = new POIFSFileSystem();

            //get the root directory
            DirectoryEntry dir = fs.Root;

            //create a entry of DocumentSummaryInformation
            DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
            dsi.Company = "NPOI Team";
            CustomProperties customProperties = dsi.CustomProperties;
            if (customProperties == null)
                customProperties = new CustomProperties();
            customProperties.Put("测试", "value A");
            customProperties.Put("BB", "value BB");
            customProperties.Put("CCC", "value CCC");
            dsi.CustomProperties = customProperties;
            //Write the stream data of the DocumentSummaryInformation entry to the root directory
            dsi.Write(dir, DocumentSummaryInformation.DEFAULT_STREAM_NAME);

            HSSFWorkbook workBook = new HSSFWorkbook();
            workBook.DocumentSummaryInformation = dsi;
            workBook.CreateSheet("aaa");//至少一个sheet 否则会报错

            downLoadFile_HSSF(workBook, "CreateCustomProperties.xls");
        }

        //创建POIFS文件
        protected void lnkbtn_CreatePOIFS_Click(object sender, EventArgs e)
        {
            POIFSFileSystem fs = new POIFSFileSystem();

            //get the root directory
            DirectoryEntry dir = fs.Root;
            //create a document entry
            dir.CreateDocument("Foo", new MemoryStream(new byte[] { 0x01, 0x02, 0x03 }));

            //create a folder
            dir.CreateDirectory("Hello");

            //create a POIFS file called Foo.poifs
            FileStream output = new FileStream("D:\\Foo.poifs", FileMode.OpenOrCreate);
            fs.WriteFileSystem(output);
            output.Close();
        }

        //POIFS文件的创建和性能
        protected void lnkbtn_CreatePOIFSFileWithPropertie_Click(object sender, EventArgs e)
        {
            POIFSFileSystem fs = new POIFSFileSystem();

            //get the root directory
            DirectoryEntry dir = fs.Root;

            //create a entry of DocumentSummaryInformation
            DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
            dsi.Company = "NPOI Team";
            //Write the stream data of the DocumentSummaryInformation entry to the root directory
            dsi.Write(dir, DocumentSummaryInformation.DEFAULT_STREAM_NAME);

            //create a entry of SummaryInformation
            SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
            si.Subject = "NPOI SDK Example";
            //Write the stream data of the SummaryInformation entry to the root directory
            si.Write(dir, SummaryInformation.DEFAULT_STREAM_NAME);

            HSSFWorkbook workBook = new HSSFWorkbook();
            workBook.DocumentSummaryInformation = dsi;
            workBook.CreateSheet("asd");
            downLoadFile_HSSF(workBook, "CreatePOIFSFileWithPropertie.xls");
        }

        //读取DB文件
        protected void lnkbtn_ReadThumbsDB_Click(object sender, EventArgs e)
        {
            FileStream stream = new FileStream(Server.MapPath("DB/thumbs.db"), FileMode.Open, FileAccess.Read);
            POIFSFileSystem poifs = new POIFSFileSystem(stream);
            var entries = poifs.Root.Entries;

            //POIFSDocumentReader catalogdr = poifs.CreatePOIFSDocumentReader("Catalog");
            //byte[] b1=new byte[catalogdr.Length-4];
            //catalogdr.Read(b1,4,b1.Length);
            //Dictionary<string, string> indexList = new Dictionary<string, string>();
            //for (int j = 0; j < b1.Length; j++)
            //{ 
            //    if(b1[0]
            //}

            while (entries.MoveNext())
            {
                DocumentNode entry = entries.Current as DocumentNode;
                DocumentInputStream dr = poifs.CreateDocumentInputStream(entry.Name);

                if (entry.Name.ToLower() == "catalog")
                    continue;

                byte[] buffer = new byte[dr.Length];
                dr.Read(buffer);
                int startpos = 0;

                //detect jfif header
                for (int i = 3; i < buffer.Length; i++)
                {
                    if (buffer[i - 3] == 0xFF
                        && buffer[i - 2] == 0xD8
                        && buffer[i - 1] == 0xFF
                        && buffer[i] == 0xE0)
                    {
                        startpos = i - 3;
                        break;
                    }
                }
                if (startpos == 0)
                    continue;

                FileStream jpeg = File.Create("D:\\" + entry.Name + ".jpeg");
                jpeg.Write(buffer, startpos, buffer.Length - startpos);
                jpeg.Close();
            }
            stream.Close();
        }

        #endregion

        #region Excel 2007

        //带有HyperLink的Excel
        protected void lnkbtnExcel_2007_HyperLink_Click(object sender, EventArgs e)
        {
            IWorkbook workBook = new XSSFWorkbook();

            ////cell style for hyperlinks
            ////by default hyperlinks are blue and underlined
            ICellStyle hlink_style = workBook.CreateCellStyle();
            IFont hlink_font = workBook.CreateFont();
            hlink_font.Underline = FontUnderlineType.Single;
            hlink_font.Color = HSSFColor.Blue.Index;
            hlink_style.SetFont(hlink_font);

            ICell cell;
            ISheet sheet = workBook.CreateSheet("Hyperlinks");

            //URL
            cell = sheet.CreateRow(0).CreateCell(0);
            cell.SetCellValue("URL Link");
            XSSFHyperlink link = new XSSFHyperlink(HyperlinkType.Url);
            link.Address = ("http://poi.apache.org/");
            cell.Hyperlink = (link);
            cell.CellStyle = (hlink_style);

            //link to a file in the current directory
            cell = sheet.CreateRow(1).CreateCell(0);
            cell.SetCellValue("File Link");
            link = new XSSFHyperlink(HyperlinkType.File);
            link.Address = ("link1.xls");
            cell.Hyperlink = (link);
            cell.CellStyle = (hlink_style);

            //e-mail link
            cell = sheet.CreateRow(2).CreateCell(0);
            cell.SetCellValue("Email Link");
            link = new XSSFHyperlink(HyperlinkType.Email);
            //note, if subject contains white spaces, make sure they are url-encoded
            link.Address = ("mailto:poi@apache.org?subject=Hyperlinks");
            cell.Hyperlink = (link);
            cell.CellStyle = (hlink_style);

            //link to a place in this workbook

            //Create a target sheet and cell
            ISheet sheet2 = workBook.CreateSheet("Target ISheet");
            sheet2.CreateRow(0).CreateCell(0).SetCellValue("Target ICell");

            cell = sheet.CreateRow(3).CreateCell(0);
            cell.SetCellValue("Worksheet Link");
            link = new XSSFHyperlink(HyperlinkType.Document);
            link.Address = ("'Target ISheet'!A1");
            cell.Hyperlink = (link);
            cell.CellStyle = (hlink_style);

            downLoadFile_XSSF(workBook, "Excel_2007_HyperLink.xlsx");
        }

        //带有字体应用的Excel
        protected void lnkbtnExcel_2007_Font_Click(object sender, EventArgs e)
        {
            IWorkbook workBook = new XSSFWorkbook();
            ISheet sheet1 = workBook.CreateSheet("Sheet1");

            //font style1: underlined, italic, red color, fontsize=20
            IFont font1 = workBook.CreateFont();
            font1.Color = IndexedColors.Red.Index;
            font1.IsItalic = true;
            font1.Underline = FontUnderlineType.Double;
            font1.FontHeightInPoints = 20;

            //bind font with style 1
            ICellStyle style1 = workBook.CreateCellStyle();
            style1.SetFont(font1);

            //font style2: strikeout line, green color, fontsize=15, fontname='宋体'
            IFont font2 = workBook.CreateFont();
            font2.Color = IndexedColors.OliveGreen.Index;
            font2.IsStrikeout = true;
            font2.FontHeightInPoints = 15;
            font2.FontName = "宋体";

            //bind font with style 2
            ICellStyle style2 = workBook.CreateCellStyle();
            style2.SetFont(font2);

            //apply font styles
            ICell cell1 = sheet1.CreateRow(1).CreateCell(1);
            cell1.SetCellValue("Hello World!");
            cell1.CellStyle = style1;
            ICell cell2 = sheet1.CreateRow(3).CreateCell(1);
            cell2.SetCellValue("早上好！");
            cell2.CellStyle = style2;

            ////cell with rich text 
            ICell cell3 = sheet1.CreateRow(5).CreateCell(1);
            XSSFRichTextString richtext = new XSSFRichTextString("Microsoft OfficeTM");

            //apply font to "Microsoft Office"
            IFont font4 = workBook.CreateFont();
            font4.FontHeightInPoints = 12;
            richtext.ApplyFont(0, 16, font4);
            //apply font to "TM"
            IFont font3 = workBook.CreateFont();
            font3.TypeOffset = FontSuperScript.Super;
            font3.IsItalic = true;
            font3.Color = IndexedColors.Blue.Index;
            font3.FontHeightInPoints = 8;
            richtext.ApplyFont(16, 18, font3);

            cell3.SetCellValue(richtext);

            downLoadFile_XSSF(workBook, "Excel_2007_Font.xlsx");
        }

        //带有表格的Excel(失败！！！)
        protected void lnkbtnExcel_2007_ApplyTable_Click(object sender, EventArgs e)
        {
            IWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet1 = (XSSFSheet)workbook.CreateSheet("Sheet1");
            sheet1.CreateRow(0).CreateCell(0).SetCellValue("This is a Sample");
            int x = 1;
            for (int i = 1; i <= 15; i++)
            {
                IRow row = sheet1.CreateRow(i);
                for (int j = 0; j < 15; j++)
                {
                    row.CreateCell(j).SetCellValue(x++);
                }
            }
            XSSFTable table = sheet1.CreateTable();
            table.Name = "Tabella1";
            table.DisplayName = "Tabella1";

            downLoadFile_XSSF(workbook, "Excel_2007_ApplyTable.xlsx");
        }

        //带有1W行数据Excel导出实验
        protected void lbkbtnExcel_2007_BigGrid_Click(object sender, EventArgs e)
        {
            IWorkbook workbook = new XSSFWorkbook();
            ISheet worksheet = workbook.CreateSheet("Sheet1");

            for (int rownum = 0; rownum < 10000; rownum++)
            {
                IRow row = worksheet.CreateRow(rownum);
                for (int celnum = 0; celnum < 20; celnum++)
                {
                    ICell Cell = row.CreateCell(celnum);
                    Cell.SetCellValue("Cell: Row-" + rownum + ";CellNo:" + celnum);
                }
            }
            downLoadFile_XSSF(workbook, "Excel_2007_BigGrid.xlsx");
        }

        //带有边框样式的Excel
        protected void lnkbtnExcel_2007_BorderStyles_Click(object sender, EventArgs e)
        {
            IWorkbook workbook = new XSSFWorkbook();
            ISheet sheet = workbook.CreateSheet("Sheet A1");
            IRow row = sheet.CreateRow(1);
            // Create a cell and put a value in it.
            ICell cell = row.CreateCell(1);
            cell.SetCellValue(4);

            // Style the cell with borders all around.
            ICellStyle style = workbook.CreateCellStyle();
            style.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            style.BottomBorderColor = IndexedColors.Black.Index;
            style.BorderLeft = NPOI.SS.UserModel.BorderStyle.DashDotDot;
            style.LeftBorderColor = IndexedColors.Green.Index;
            style.BorderRight = NPOI.SS.UserModel.BorderStyle.Hair;
            style.RightBorderColor = IndexedColors.Blue.Index;
            style.BorderTop = NPOI.SS.UserModel.BorderStyle.MediumDashed;
            style.TopBorderColor = IndexedColors.Orange.Index;

            //create border diagonal
            style.BorderDiagonalLineStyle = NPOI.SS.UserModel.BorderStyle.Medium; //this property must be set before BorderDiagonal and BorderDiagonalColor
            style.BorderDiagonal = BorderDiagonal.Forward;
            style.BorderDiagonalColor = IndexedColors.Gold.Index;

            cell.CellStyle = style;
            // Create a cell and put a value in it.
            ICell cell2 = row.CreateCell(2);
            cell2.SetCellValue(5);
            ICellStyle style2 = workbook.CreateCellStyle();
            style2.BorderDiagonalLineStyle = NPOI.SS.UserModel.BorderStyle.Medium;
            style2.BorderDiagonal = BorderDiagonal.Backward;
            style2.BorderDiagonalColor = IndexedColors.Red.Index;
            cell2.CellStyle = style2;

            downLoadFile_XSSF(workbook, "Excel_2007_BorderStyles.xlsx");
        }

        //带有颜色矩阵的Excel
        protected void lnkbtnExcel_2007_ColorfulMatrix_Click(object sender, EventArgs e)
        {
            IWorkbook workbook = new XSSFWorkbook();
            ISheet sheet1 = workbook.CreateSheet("Sheet1");
            int x = 1;
            for (int i = 0; i < 15; i++)
            {
                IRow row = sheet1.CreateRow(i);
                for (int j = 0; j < 15; j++)
                {
                    ICell cell = row.CreateCell(j);
                    if (x % 2 == 0)
                    {
                        //fill background with blue
                        ICellStyle style1 = workbook.CreateCellStyle();
                        style1.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Blue.Index2;
                        style1.FillPattern = FillPattern.SolidForeground;
                        cell.CellStyle = style1;
                    }
                    else
                    {
                        //fill background with yellow
                        ICellStyle style1 = workbook.CreateCellStyle();
                        style1.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Yellow.Index2;
                        style1.FillPattern = FillPattern.SolidForeground;
                        cell.CellStyle = style1;
                    }
                    x++;
                }
            }

            downLoadFile_XSSF(workbook, "Excel_2007_ColorfulMatrix.xlsx");
        }

        #region  带有格式规则的Excel

        protected void lnkbtnExcel_2007_ConditionalFormats_Click(object sender, EventArgs e)
        {
            IWorkbook workBook = new XSSFWorkbook();

            SameCell(workBook.CreateSheet("Same Cell"));
            MultiCell(workBook.CreateSheet("MultiCell"));
            Errors(workBook.CreateSheet("Errors"));
            HideDupplicates(workBook.CreateSheet("Hide Dups"));
            FormatDuplicates(workBook.CreateSheet("Duplicates"));
            InList(workBook.CreateSheet("In List"));
            Expiry(workBook.CreateSheet("Expiry"));
            ShadeAlt(workBook.CreateSheet("Shade Alt"));
            ShadeBands(workBook.CreateSheet("Shade Bands"));

            downLoadFile_XSSF(workBook, "Excel_2007_ConditionalFormats.xlsx");
        }

        static void SameCell(ISheet sheet)
        {
            sheet.CreateRow(0).CreateCell(0).SetCellValue(84);
            sheet.CreateRow(1).CreateCell(0).SetCellValue(74);
            sheet.CreateRow(2).CreateCell(0).SetCellValue(50);
            sheet.CreateRow(3).CreateCell(0).SetCellValue(51);
            sheet.CreateRow(4).CreateCell(0).SetCellValue(49);
            sheet.CreateRow(5).CreateCell(0).SetCellValue(41);

            ISheetConditionalFormatting sheetCF = sheet.SheetConditionalFormatting;

            // Condition 1: Cell Value Is   greater than  70   (Blue Fill)
            IConditionalFormattingRule rule1 = sheetCF.CreateConditionalFormattingRule(ComparisonOperator.GreaterThan, "70");
            IPatternFormatting fill1 = rule1.CreatePatternFormatting();
            fill1.FillBackgroundColor = (IndexedColors.Blue.Index);
            fill1.FillPattern = (short)FillPattern.SolidForeground;

            // Condition 2: Cell Value Is  less than      50   (Green Fill)
            IConditionalFormattingRule rule2 = sheetCF.CreateConditionalFormattingRule(ComparisonOperator.LessThan, "50");
            IPatternFormatting fill2 = rule2.CreatePatternFormatting();
            fill2.FillBackgroundColor = (IndexedColors.Green.Index);
            fill2.FillPattern = (short)FillPattern.SolidForeground;

            CellRangeAddress[] regions = {
                CellRangeAddress.ValueOf("A1:A6")
        };

            sheetCF.AddConditionalFormatting(regions, rule1, rule2);

            sheet.GetRow(0).CreateCell(2).SetCellValue("<== Condition 1: Cell Value is greater than 70 (Blue Fill)");
            sheet.GetRow(4).CreateCell(2).SetCellValue("<== Condition 2: Cell Value is less than 50 (Green Fill)");
        }

        /**
         * Highlight multiple cells based on a formula
         */
        static void MultiCell(ISheet sheet)
        {
            // header row
            IRow row0 = sheet.CreateRow(0);
            row0.CreateCell(0).SetCellValue("Units");
            row0.CreateCell(1).SetCellValue("Cost");
            row0.CreateCell(2).SetCellValue("Total");

            IRow row1 = sheet.CreateRow(1);
            row1.CreateCell(0).SetCellValue(71);
            row1.CreateCell(1).SetCellValue(29);
            row1.CreateCell(2).SetCellValue(2059);

            IRow row2 = sheet.CreateRow(2);
            row2.CreateCell(0).SetCellValue(85);
            row2.CreateCell(1).SetCellValue(29);
            row2.CreateCell(2).SetCellValue(2059);

            IRow row3 = sheet.CreateRow(3);
            row3.CreateCell(0).SetCellValue(71);
            row3.CreateCell(1).SetCellValue(29);
            row3.CreateCell(2).SetCellValue(2059);

            ISheetConditionalFormatting sheetCF = sheet.SheetConditionalFormatting;

            // Condition 1: Formula Is   =$B2>75   (Blue Fill)
            IConditionalFormattingRule rule1 = sheetCF.CreateConditionalFormattingRule("$A2>75");
            IPatternFormatting fill1 = rule1.CreatePatternFormatting();
            fill1.FillBackgroundColor = (IndexedColors.Blue.Index);
            fill1.FillPattern = ((short)FillPattern.SolidForeground);

            CellRangeAddress[] regions = {
                CellRangeAddress.ValueOf("A2:C4")
        };

            sheetCF.AddConditionalFormatting(regions, rule1);

            sheet.GetRow(2).CreateCell(4).SetCellValue("<== Condition 1: Formula is =$B2>75   (Blue Fill)");
        }

        /**
         *  Use Excel conditional formatting to check for errors,
         *  and change the font colour to match the cell colour.
         *  In this example, if formula result is  #DIV/0! then it will have white font colour.
         */
        static void Errors(ISheet sheet)
        {
            sheet.CreateRow(0).CreateCell(0).SetCellValue(84);
            sheet.CreateRow(1).CreateCell(0).SetCellValue(0);
            sheet.CreateRow(2).CreateCell(0).SetCellFormula("ROUND(A1/A2,0)");
            sheet.CreateRow(3).CreateCell(0).SetCellValue(0);
            sheet.CreateRow(4).CreateCell(0).SetCellFormula("ROUND(A6/A4,0)");
            sheet.CreateRow(5).CreateCell(0).SetCellValue(41);

            ISheetConditionalFormatting sheetCF = sheet.SheetConditionalFormatting;

            // Condition 1: Formula Is   =ISERROR(C2)   (White Font)
            IConditionalFormattingRule rule1 = sheetCF.CreateConditionalFormattingRule("ISERROR(A1)");
            IFontFormatting font = rule1.CreateFontFormatting();
            font.FontColorIndex = (IndexedColors.White.Index);

            CellRangeAddress[] regions = {
                CellRangeAddress.ValueOf("A1:A6")
        };

            sheetCF.AddConditionalFormatting(regions, rule1);

            sheet.GetRow(2).CreateCell(1).SetCellValue("<== The error in this cell is hidden. Condition: Formula is   =ISERROR(C2)   (White Font)");
            sheet.GetRow(4).CreateCell(1).SetCellValue("<== The error in this cell is hidden. Condition: Formula is   =ISERROR(C2)   (White Font)");
        }

        /**
         * Use Excel conditional formatting to hide the duplicate values,
         * and make the list easier to read. In this example, when the table is sorted by Region,
         * the second (and subsequent) occurences of each region name will have white font colour.
         */
        static void HideDupplicates(ISheet sheet)
        {
            sheet.CreateRow(0).CreateCell(0).SetCellValue("City");
            sheet.CreateRow(1).CreateCell(0).SetCellValue("Boston");
            sheet.CreateRow(2).CreateCell(0).SetCellValue("Boston");
            sheet.CreateRow(3).CreateCell(0).SetCellValue("Chicago");
            sheet.CreateRow(4).CreateCell(0).SetCellValue("Chicago");
            sheet.CreateRow(5).CreateCell(0).SetCellValue("New York");

            ISheetConditionalFormatting sheetCF = sheet.SheetConditionalFormatting;

            // Condition 1: Formula Is   =A2=A1   (White Font)
            IConditionalFormattingRule rule1 = sheetCF.CreateConditionalFormattingRule("A2=A1");
            IFontFormatting font = rule1.CreateFontFormatting();
            font.FontColorIndex = IndexedColors.White.Index;

            CellRangeAddress[] regions = {
                CellRangeAddress.ValueOf("A2:A6")
            };

            sheetCF.AddConditionalFormatting(regions, rule1);

            sheet.GetRow(1).CreateCell(1).SetCellValue("<== the second (and subsequent) " +
                    "occurences of each region name will have white font colour.  " +
                    "Condition: Formula Is   =A2=A1   (White Font)");
        }

        /**
         * Use Excel conditional formatting to highlight duplicate entries in a column.
         */
        static void FormatDuplicates(ISheet sheet)
        {
            sheet.CreateRow(0).CreateCell(0).SetCellValue("Code");
            sheet.CreateRow(1).CreateCell(0).SetCellValue(4);
            sheet.CreateRow(2).CreateCell(0).SetCellValue(3);
            sheet.CreateRow(3).CreateCell(0).SetCellValue(6);
            sheet.CreateRow(4).CreateCell(0).SetCellValue(3);
            sheet.CreateRow(5).CreateCell(0).SetCellValue(5);
            sheet.CreateRow(6).CreateCell(0).SetCellValue(8);
            sheet.CreateRow(7).CreateCell(0).SetCellValue(0);
            sheet.CreateRow(8).CreateCell(0).SetCellValue(2);
            sheet.CreateRow(9).CreateCell(0).SetCellValue(8);
            sheet.CreateRow(10).CreateCell(0).SetCellValue(6);

            ISheetConditionalFormatting sheetCF = sheet.SheetConditionalFormatting;

            // Condition 1: Formula Is   =A2=A1   (White Font)
            IConditionalFormattingRule rule1 = sheetCF.CreateConditionalFormattingRule("COUNTIF($A$2:$A$11,A2)>1");
            IFontFormatting font = rule1.CreateFontFormatting();
            font.SetFontStyle(false, true);
            font.FontColorIndex = (IndexedColors.Blue.Index);

            CellRangeAddress[] regions = {
                CellRangeAddress.ValueOf("A2:A11")
        };

            sheetCF.AddConditionalFormatting(regions, rule1);

            sheet.GetRow(2).CreateCell(1).SetCellValue("<== Duplicates numbers in the column are highlighted.  " +
                    "Condition: Formula Is =COUNTIF($A$2:$A$11,A2)>1   (Blue Font)");
        }

        /**
         * Use Excel conditional formatting to highlight items that are in a list on the worksheet.
         */
        static void InList(ISheet sheet)
        {
            sheet.CreateRow(0).CreateCell(0).SetCellValue("Codes");
            sheet.CreateRow(1).CreateCell(0).SetCellValue("AA");
            sheet.CreateRow(2).CreateCell(0).SetCellValue("BB");
            sheet.CreateRow(3).CreateCell(0).SetCellValue("GG");
            sheet.CreateRow(4).CreateCell(0).SetCellValue("AA");
            sheet.CreateRow(5).CreateCell(0).SetCellValue("FF");
            sheet.CreateRow(6).CreateCell(0).SetCellValue("XX");
            sheet.CreateRow(7).CreateCell(0).SetCellValue("CC");

            sheet.GetRow(0).CreateCell(2).SetCellValue("Valid");
            sheet.GetRow(1).CreateCell(2).SetCellValue("AA");
            sheet.GetRow(2).CreateCell(2).SetCellValue("BB");
            sheet.GetRow(3).CreateCell(2).SetCellValue("CC");

            ISheetConditionalFormatting sheetCF = sheet.SheetConditionalFormatting;

            // Condition 1: Formula Is   =A2=A1   (White Font)
            IConditionalFormattingRule rule1 = sheetCF.CreateConditionalFormattingRule("COUNTIF($C$2:$C$4,A2)");
            IPatternFormatting fill1 = rule1.CreatePatternFormatting();
            fill1.FillBackgroundColor = (IndexedColors.LightBlue.Index);
            fill1.FillPattern = ((short)FillPattern.SolidForeground);

            CellRangeAddress[] regions = {
                CellRangeAddress.ValueOf("A2:A8")
        };

            sheetCF.AddConditionalFormatting(regions, rule1);

            sheet.GetRow(2).CreateCell(3).SetCellValue("<== Use Excel conditional formatting to highlight items that are in a list on the worksheet");
        }

        /**
         *  Use Excel conditional formatting to highlight payments that are due in the next thirty days.
         *  In this example, Due dates are entered in cells A2:A4.
         */
        static void Expiry(ISheet sheet)
        {
            ICellStyle style = sheet.Workbook.CreateCellStyle();
            style.DataFormat = (short)BuiltinFormats.GetBuiltinFormat("d-mmm");

            sheet.CreateRow(0).CreateCell(0).SetCellValue("Date");
            sheet.CreateRow(1).CreateCell(0).SetCellFormula("TODAY()+29");
            sheet.CreateRow(2).CreateCell(0).SetCellFormula("A2+1");
            sheet.CreateRow(3).CreateCell(0).SetCellFormula("A3+1");

            for (int rownum = 1; rownum <= 3; rownum++) sheet.GetRow(rownum).GetCell(0).CellStyle = style;

            ISheetConditionalFormatting sheetCF = sheet.SheetConditionalFormatting;

            // Condition 1: Formula Is   =A2=A1   (White Font)
            IConditionalFormattingRule rule1 = sheetCF.CreateConditionalFormattingRule("AND(A2-TODAY()>=0,A2-TODAY()<=30)");
            IFontFormatting font = rule1.CreateFontFormatting();
            font.SetFontStyle(false, true);
            font.FontColorIndex = IndexedColors.Blue.Index;

            CellRangeAddress[] regions = {
                CellRangeAddress.ValueOf("A2:A4")
        };

            sheetCF.AddConditionalFormatting(regions, rule1);

            sheet.GetRow(0).CreateCell(1).SetCellValue("Dates within the next 30 days are highlighted");
        }

        /**
         * Use Excel conditional formatting to shade alternating rows on the worksheet
         */
        static void ShadeAlt(ISheet sheet)
        {
            ISheetConditionalFormatting sheetCF = sheet.SheetConditionalFormatting;

            // Condition 1: Formula Is   =A2=A1   (White Font)
            IConditionalFormattingRule rule1 = sheetCF.CreateConditionalFormattingRule("MOD(ROW(),2)");
            IPatternFormatting fill1 = rule1.CreatePatternFormatting();
            fill1.FillBackgroundColor = (IndexedColors.LightGreen.Index);
            fill1.FillPattern = ((short)FillPattern.SolidForeground);

            CellRangeAddress[] regions = {
                CellRangeAddress.ValueOf("A1:Z100")
        };

            sheetCF.AddConditionalFormatting(regions, rule1);

            sheet.CreateRow(0).CreateCell(1).SetCellValue("Shade Alternating Rows");
            sheet.CreateRow(1).CreateCell(1).SetCellValue("Condition: Formula Is  =MOD(ROW(),2)   (Light Green Fill)");
        }

        /**
         * You can use Excel conditional formatting to shade bands of rows on the worksheet. 
         * In this example, 3 rows are shaded light grey, and 3 are left with no shading.
         * In the MOD function, the total number of rows in the set of banded rows (6) is entered.
         */
        static void ShadeBands(ISheet sheet)
        {
            ISheetConditionalFormatting sheetCF = sheet.SheetConditionalFormatting;

            IConditionalFormattingRule rule1 = sheetCF.CreateConditionalFormattingRule("MOD(ROW(),6)<3");
            IPatternFormatting fill1 = rule1.CreatePatternFormatting();
            fill1.FillBackgroundColor = (IndexedColors.Grey25Percent.Index);
            fill1.FillPattern = ((short)FillPattern.SolidForeground);

            CellRangeAddress[] regions = {
                CellRangeAddress.ValueOf("A1:Z100")
        };

            sheetCF.AddConditionalFormatting(regions, rule1);

            sheet.CreateRow(0).CreateCell(1).SetCellValue("Shade Bands of Rows");
            sheet.CreateRow(1).CreateCell(1).SetCellValue("Condition: Formula is  =MOD(ROW(),6)<2   (Light Grey Fill)");
        }

        #endregion

        //带有批注的Excel
        protected void lnkbtnExcel_2007_CreateComment_Click(object sender, EventArgs e)
        {
            IWorkbook workbook = new XSSFWorkbook();
            ISheet sheet = workbook.CreateSheet("some comments");

            // Create the drawing patriarch. This is the top level container for all shapes including cell comments.
            IDrawing patr = sheet.CreateDrawingPatriarch();

            //Create a cell in row 3
            ICell cell1 = sheet.CreateRow(3).CreateCell(1);
            cell1.SetCellValue(new XSSFRichTextString("Hello, World"));

            //anchor defines size and position of the comment in worksheet
            IComment comment1 = patr.CreateCellComment(new XSSFClientAnchor(0, 0, 0, 0, 4, 2, 6, 5));

            // set text in the comment
            comment1.String = (new XSSFRichTextString("We can set comments in POI"));

            //set comment author.
            //you can see it in the status bar when moving mouse over the commented cell
            comment1.Author = ("Apache Software Foundation");

            // The first way to assign comment to a cell is via HSSFCell.SetCellComment method
            cell1.CellComment = (comment1);

            //Create another cell in row 6
            ICell cell2 = sheet.CreateRow(6).CreateCell(1);
            cell2.SetCellValue(36.6);


            IComment comment2 = patr.CreateCellComment(new XSSFClientAnchor(0, 0, 0, 0, 4, 8, 6, 11));
            //modify background color of the comment
            //comment2.SetFillColor(204, 236, 255);

            XSSFRichTextString str = new XSSFRichTextString("Normal body temperature");

            //apply custom font to the text in the comment
            IFont font = workbook.CreateFont();
            font.FontName = ("Arial");
            font.FontHeightInPoints = 10;
            font.Boldweight = (short)FontBoldWeight.Bold;
            font.Color = HSSFColor.Red.Index;
            str.ApplyFont(font);

            comment2.String = str;
            comment2.Visible = true; //by default comments are hidden. This one is always visible.

            comment2.Author = "Bill Gates";

            /**
             * The second way to assign comment to a cell is to implicitly specify its row and column.
             * Note, it is possible to set row and column of a non-existing cell.
             * It works, the commnet is visible.
             */
            comment2.Row = 6;
            comment2.Column = 1;

            downLoadFile_XSSF(workbook, "Excel_2007_CreateComment.xlsx");
        }

        //创建用户自定义属性的Excel
        protected void lnkbtExcel_2007_CreateCustomProperties_Click(object sender, EventArgs e)
        {
            XSSFWorkbook workbook = new XSSFWorkbook();
            ISheet sheet1 = workbook.CreateSheet("Sheet1");

            NPOI.POIXMLProperties props = workbook.GetProperties();
            props.CoreProperties.Creator = "NPOI 2.0.5";
            props.CoreProperties.Created = DateTime.Now;
            if (!props.CustomProperties.Contains("NPOI Team"))
                props.CustomProperties.AddProperty("NPOI Team", "Hello World!");

            downLoadFile_XSSF(workbook, "Excel_2007_CreateCustomProperties.xlsx");
        }

        //带有页眉页脚的Excel
        protected void lnkbtnExcel_2007_CreateHeaderFooter_Click(object sender, EventArgs e)
        {
            IWorkbook workbook = new XSSFWorkbook();
            ISheet s1 = workbook.CreateSheet("Sheet1");
            s1.CreateRow(0).CreateCell(1).SetCellValue(123);

            //set header text
            s1.Header.Left = HSSFHeader.Page;   //Page is a static property of HSSFHeader and HSSFFooter
            s1.Header.Center = "This is a test sheet";
            //set footer text
            s1.Footer.Left = "Copyright NPOI Team";
            s1.Footer.Right = "created by Tony Qu（瞿杰）";

            downLoadFile_XSSF(workbook, "Excel_2007_CreateHeaderFooter.xlsx");
        }

        #region 带有日期格式的Excel

        protected void lnkbtnExcel_2007_DataFormats_Click(object sender, EventArgs e)
        {
            IWorkbook workbook = new XSSFWorkbook();

            ISheet sheet = workbook.CreateSheet("Sheet1");
            //increase the width of Column A
            sheet.SetColumnWidth(0, 5000);
            //create the format instance
            IDataFormat format = workbook.CreateDataFormat();

            // Create a row and put some cells in it. Rows are 0 based.
            ICell cell = sheet.CreateRow(0).CreateCell(0);
            //number format with 2 digits after the decimal point - "1.20"
            SetValueAndFormat(workbook, cell, 1.2, HSSFDataFormat.GetBuiltinFormat("0.00"));

            //RMB currency format with comma    -   "¥20,000"
            ICell cell2 = sheet.CreateRow(1).CreateCell(0);
            SetValueAndFormat(workbook, cell2, 20000, format.GetFormat("¥#,##0"));

            //scentific number format   -   "3.15E+00"
            ICell cell3 = sheet.CreateRow(2).CreateCell(0);
            SetValueAndFormat(workbook, cell3, 3.151234, format.GetFormat("0.00E+00"));

            //percent format, 2 digits after the decimal point    -  "99.33%"
            ICell cell4 = sheet.CreateRow(3).CreateCell(0);
            SetValueAndFormat(workbook, cell4, 0.99333, format.GetFormat("0.00%"));

            //phone number format - "021-65881234"
            ICell cell5 = sheet.CreateRow(4).CreateCell(0);
            SetValueAndFormat(workbook, cell5, 02165881234, format.GetFormat("000-00000000"));

            //Chinese capitalized character number - 壹贰叁 元
            ICell cell6 = sheet.CreateRow(5).CreateCell(0);
            SetValueAndFormat(workbook, cell6, 123, format.GetFormat("[DbNum2][$-804]0 元"));

            //Chinese date string
            ICell cell7 = sheet.CreateRow(6).CreateCell(0);
            SetValueAndFormat(workbook, cell7, new DateTime(2004, 5, 6), format.GetFormat("yyyy年m月d日"));
            cell7.SetCellValue(new DateTime(2004, 5, 6));

            //Chinese date string
            ICell cell8 = sheet.CreateRow(7).CreateCell(0);
            SetValueAndFormat(workbook, cell8, new DateTime(2005, 11, 6), format.GetFormat("yyyy年m月d日"));

            //formula value with datetime style 
            ICell cell9 = sheet.CreateRow(8).CreateCell(0);
            cell9.CellFormula = "DateValue(\"2005-11-11\")+TIMEVALUE(\"11:11:11\")";
            ICellStyle cellStyle9 = workbook.CreateCellStyle();
            cellStyle9.DataFormat = HSSFDataFormat.GetBuiltinFormat("m/d/yy h:mm");
            cell9.CellStyle = cellStyle9;

            //display current time
            ICell cell10 = sheet.CreateRow(9).CreateCell(0);
            SetValueAndFormat(workbook, cell10, DateTime.Now, format.GetFormat("[$-409]h:mm:ss AM/PM;@"));

            downLoadFile_XSSF(workbook, "Excel_2007_DataFormats.xlsx");
        }

        private void SetValueAndFormat(IWorkbook workbook, ICell cell, int value, short formatId)
        {
            cell.SetCellValue(value);
            ICellStyle cellStyle = workbook.CreateCellStyle();
            cellStyle.DataFormat = formatId;
            cell.CellStyle = cellStyle;
        }
        private void SetValueAndFormat(IWorkbook workbook, ICell cell, double value, short formatId)
        {
            cell.SetCellValue(value);
            ICellStyle cellStyle = workbook.CreateCellStyle();
            cellStyle.DataFormat = formatId;
            cell.CellStyle = cellStyle;
        }
        private void SetValueAndFormat(IWorkbook workbook, ICell cell, DateTime value, short formatId)
        {
            //set value for the cell
            if (value != null)
                cell.SetCellValue(value);

            ICellStyle cellStyle = workbook.CreateCellStyle();
            cellStyle.DataFormat = formatId;
            cell.CellStyle = cellStyle;
        }

        #endregion

        //带有填充背景的Excel
        protected void lnkbtnExcel_2007_FillBackground_Click(object sender, EventArgs e)
        {
            IWorkbook workbook = new XSSFWorkbook();

            ISheet sheet1 = workbook.CreateSheet("Sheet1");

            //fill background
            ICellStyle style1 = workbook.CreateCellStyle();
            style1.FillForegroundColor = IndexedColors.Blue.Index;
            style1.FillPattern = FillPattern.BigSpots;
            style1.FillBackgroundColor = IndexedColors.Pink.Index;
            sheet1.CreateRow(0).CreateCell(0).CellStyle = style1;

            //fill background
            ICellStyle style2 = workbook.CreateCellStyle();
            style2.FillForegroundColor = IndexedColors.Yellow.Index;
            style2.FillPattern = FillPattern.AltBars;
            style2.FillBackgroundColor = IndexedColors.Rose.Index;
            sheet1.CreateRow(1).CreateCell(0).CellStyle = style2;

            //fill background
            ICellStyle style3 = workbook.CreateCellStyle();
            style3.FillForegroundColor = IndexedColors.Lime.Index;
            style3.FillPattern = FillPattern.LessDots;
            style3.FillBackgroundColor = IndexedColors.LightGreen.Index;
            sheet1.CreateRow(2).CreateCell(0).CellStyle = style3;

            //fill background
            ICellStyle style4 = workbook.CreateCellStyle();
            style4.FillForegroundColor = IndexedColors.Yellow.Index;
            style4.FillPattern = FillPattern.LeastDots;
            style4.FillBackgroundColor = IndexedColors.Rose.Index;
            sheet1.CreateRow(3).CreateCell(0).CellStyle = style4;

            //fill background
            ICellStyle style5 = workbook.CreateCellStyle();
            style5.FillForegroundColor = IndexedColors.LightBlue.Index;
            style5.FillPattern = FillPattern.Bricks;
            style5.FillBackgroundColor = IndexedColors.Plum.Index;
            sheet1.CreateRow(4).CreateCell(0).CellStyle = style5;

            //fill background
            ICellStyle style6 = workbook.CreateCellStyle();
            style6.FillForegroundColor = IndexedColors.SeaGreen.Index;
            style6.FillPattern = FillPattern.FineDots;
            style6.FillBackgroundColor = IndexedColors.White.Index;
            sheet1.CreateRow(5).CreateCell(0).CellStyle = style6;

            //fill background
            ICellStyle style7 = workbook.CreateCellStyle();
            style7.FillForegroundColor = IndexedColors.Orange.Index;
            style7.FillPattern = FillPattern.Diamonds;
            style7.FillBackgroundColor = IndexedColors.Orchid.Index;
            sheet1.CreateRow(6).CreateCell(0).CellStyle = style7;

            //fill background
            ICellStyle style8 = workbook.CreateCellStyle();
            style8.FillForegroundColor = IndexedColors.White.Index;
            style8.FillPattern = FillPattern.Squares;
            style8.FillBackgroundColor = IndexedColors.Red.Index;
            sheet1.CreateRow(7).CreateCell(0).CellStyle = style8;

            //fill background
            ICellStyle style9 = workbook.CreateCellStyle();
            style9.FillForegroundColor = IndexedColors.RoyalBlue.Index;
            style9.FillPattern = FillPattern.SparseDots;
            style9.FillBackgroundColor = IndexedColors.Yellow.Index;
            sheet1.CreateRow(8).CreateCell(0).CellStyle = style9;

            //fill background
            ICellStyle style10 = workbook.CreateCellStyle();
            style10.FillForegroundColor = IndexedColors.RoyalBlue.Index;
            style10.FillPattern = FillPattern.ThinBackwardDiagonals;
            style10.FillBackgroundColor = IndexedColors.Yellow.Index;
            sheet1.CreateRow(9).CreateCell(0).CellStyle = style10;

            //fill background
            ICellStyle style11 = workbook.CreateCellStyle();
            style11.FillForegroundColor = IndexedColors.RoyalBlue.Index;
            style11.FillPattern = FillPattern.ThickForwardDiagonals;
            style11.FillBackgroundColor = IndexedColors.Yellow.Index;
            sheet1.CreateRow(10).CreateCell(0).CellStyle = style11;

            //fill background
            ICellStyle style12 = workbook.CreateCellStyle();
            style12.FillForegroundColor = IndexedColors.RoyalBlue.Index;
            style12.FillPattern = FillPattern.ThickHorizontalBands;
            style12.FillBackgroundColor = IndexedColors.Yellow.Index;
            sheet1.CreateRow(11).CreateCell(0).CellStyle = style12;


            //fill background
            ICellStyle style13 = workbook.CreateCellStyle();
            style13.FillForegroundColor = IndexedColors.RoyalBlue.Index;
            style13.FillPattern = FillPattern.ThickVerticalBands;
            style13.FillBackgroundColor = IndexedColors.Yellow.Index;
            sheet1.CreateRow(12).CreateCell(0).CellStyle = style13;

            //fill background
            ICellStyle style14 = workbook.CreateCellStyle();
            style14.FillForegroundColor = IndexedColors.RoyalBlue.Index;
            style14.FillPattern = FillPattern.ThickBackwardDiagonals;
            style14.FillBackgroundColor = IndexedColors.Yellow.Index;
            sheet1.CreateRow(13).CreateCell(0).CellStyle = style14;

            //fill background
            ICellStyle style15 = workbook.CreateCellStyle();
            style15.FillForegroundColor = IndexedColors.RoyalBlue.Index;
            style15.FillPattern = FillPattern.ThinForwardDiagonals;
            style15.FillBackgroundColor = IndexedColors.Yellow.Index;
            sheet1.CreateRow(14).CreateCell(0).CellStyle = style15;

            //fill background
            ICellStyle style16 = workbook.CreateCellStyle();
            style16.FillForegroundColor = IndexedColors.RoyalBlue.Index;
            style16.FillPattern = FillPattern.ThinHorizontalBands;
            style16.FillBackgroundColor = IndexedColors.Yellow.Index;
            sheet1.CreateRow(15).CreateCell(0).CellStyle = style16;

            //fill background
            ICellStyle style17 = workbook.CreateCellStyle();
            style17.FillForegroundColor = IndexedColors.RoyalBlue.Index;
            style17.FillPattern = FillPattern.ThinVerticalBands;
            style17.FillBackgroundColor = IndexedColors.Yellow.Index;
            sheet1.CreateRow(16).CreateCell(0).CellStyle = style17;

            downLoadFile_XSSF(workbook, "Excel_2007_FillBackground.xlsx");
        }

        //带有隐藏行和列的Excel
        protected void lnkbtnExcel_2007_HideColumnAndRow_Click(object sender, EventArgs e)
        {
            IWorkbook workbook = new XSSFWorkbook();
            ISheet s = workbook.CreateSheet("Sheet1");
            IRow r1 = s.CreateRow(0);
            IRow r2 = s.CreateRow(1);
            IRow r3 = s.CreateRow(2);
            IRow r4 = s.CreateRow(3);
            IRow r5 = s.CreateRow(4);

            //hide IRow 2
            r2.ZeroHeight = true;

            //hide column C
            s.SetColumnHidden(2, true);

            downLoadFile_XSSF(workbook, "Excel_2007_HideColumnAndRow.xlsx");
        }

        //带有图片的Excel
        protected void lnkbtnExcel_2007_InsertPictures_Click(object sender, EventArgs e)
        {
            IWorkbook workbook = new XSSFWorkbook();
            ISheet sheet1 = workbook.CreateSheet("PictureSheet");


            IDrawing patriarch = sheet1.CreateDrawingPatriarch();
            //create the anchor
            XSSFClientAnchor anchor = new XSSFClientAnchor(500, 200, 0, 0, 2, 2, 4, 7);
            anchor.AnchorType = 2;
            //load the picture and get the picture index in the workbook
            //first picture

            string path = Server.MapPath("Images/HumpbackWhale.jpg");
            FileStream file = new FileStream(path, FileMode.Open, FileAccess.Read);
            byte[] buffer = new byte[file.Length];
            file.Read(buffer, 0, (int)file.Length);

            int imageId = workbook.AddPicture(buffer, NPOI.SS.UserModel.PictureType.JPEG);
            XSSFPicture picture = (XSSFPicture)patriarch.CreatePicture(anchor, imageId);
            //Reset the image to the original size.
            //picture.Resize();   //Note: Resize will reset client anchor you set.
            picture.LineStyle = LineStyle.DashDotGel;

            //second picture            
            XSSFClientAnchor anchor2 = new XSSFClientAnchor(500, 200, 0, 0, 5, 10, 7, 15);
            XSSFPicture picture2 = (XSSFPicture)patriarch.CreatePicture(anchor2, imageId);
            picture.LineStyle = LineStyle.DashDotGel;

            downLoadFile_XSSF(workbook, "Excel_2007_InsertPictures.xlsx");
        }

        //带有线型图表的Excel
        protected void lnkbtnExcel_2007_LineChart_Click(object sender, EventArgs e)
        {
            IWorkbook workBook = new XSSFWorkbook();
            ISheet sheet = workBook.CreateSheet("linechart");
            int NUM_OF_ROWS = 3;
            int NUM_OF_COLUMNS = 10;

            // Create a row and put some cells in it. Rows are 0 based.
            IRow row;
            ICell cell;
            for (int rowIndex = 0; rowIndex < NUM_OF_ROWS; rowIndex++)
            {
                row = sheet.CreateRow((short)rowIndex);
                for (int colIndex = 0; colIndex < NUM_OF_COLUMNS; colIndex++)
                {
                    cell = row.CreateCell((short)colIndex);
                    cell.SetCellValue(colIndex * (rowIndex + 1));
                }
            }

            IDrawing drawing = sheet.CreateDrawingPatriarch();
            IClientAnchor anchor = drawing.CreateAnchor(0, 0, 0, 0, 0, 5, 10, 15);

            IChart chart = drawing.CreateChart(anchor);
            NPOI.SS.UserModel.Charts.IChartLegend legend = chart.GetOrCreateLegend();
            legend.Position = NPOI.SS.UserModel.Charts.LegendPosition.TopRight;

            NPOI.SS.UserModel.Charts.ILineChartData<double, double> data = chart.GetChartDataFactory().CreateLineChartData<double, double>();

            // Use a category axis for the bottom axis.
            NPOI.SS.UserModel.Charts.IChartAxis bottomAxis = chart.GetChartAxisFactory().CreateCategoryAxis(NPOI.SS.UserModel.Charts.AxisPosition.Bottom);
            NPOI.SS.UserModel.Charts.IValueAxis leftAxis = chart.GetChartAxisFactory().CreateValueAxis(NPOI.SS.UserModel.Charts.AxisPosition.Left);
            leftAxis.SetCrosses(NPOI.SS.UserModel.Charts.AxisCrosses.AutoZero);

            NPOI.SS.UserModel.Charts.IChartDataSource<double> xs = NPOI.SS.UserModel.Charts.DataSources.FromNumericCellRange(sheet, new CellRangeAddress(0, 0, 0, NUM_OF_COLUMNS - 1));
            NPOI.SS.UserModel.Charts.IChartDataSource<double> ys1 = NPOI.SS.UserModel.Charts.DataSources.FromNumericCellRange(sheet, new CellRangeAddress(1, 1, 0, NUM_OF_COLUMNS - 1));
            NPOI.SS.UserModel.Charts.IChartDataSource<double> ys2 = NPOI.SS.UserModel.Charts.DataSources.FromNumericCellRange(sheet, new CellRangeAddress(2, 2, 0, NUM_OF_COLUMNS - 1));


            var s1 = data.AddSerie(xs, ys1);
            s1.SetTitle("title1");
            var s2 = data.AddSerie(xs, ys2);
            s2.SetTitle("title2");

            chart.Plot(data, bottomAxis, leftAxis);
            downLoadFile_XSSF(workBook, "Excel_2007_LineChart.xlsx");
        }

        //带有合并单元格的Excel
        protected void lnkbtnExcel_2007_MeringCells_Click(object sender, EventArgs e)
        {
            IWorkbook workbook = new XSSFWorkbook();
            ISheet sheet = workbook.CreateSheet("PictureSheet");

            IRow row = sheet.CreateRow(1);
            ICell cell = row.CreateCell(1);
            cell.SetCellValue(new XSSFRichTextString("This is a test of merging"));

            sheet.AddMergedRegion(new CellRangeAddress(1, 1, 1, 2));

            downLoadFile_XSSF(workbook, "Excel_2007_MeringCells.xlsx");
        }

        #region 带有下拉框的Excel

        protected void lnkbtnExcel_2007_LinkedDropDownLists_Click(object sender, EventArgs e)
        {
            IWorkbook workbook = new XSSFWorkbook();
            ISheet sheet = null;
            IDataValidationHelper dvHelper = null;
            IDataValidationConstraint dvConstraint = null;
            IDataValidation validation = null;
            CellRangeAddressList addressList = null;

            sheet = workbook.CreateSheet("Linked Validations");
            BuildDataSheet(sheet);

            addressList = new CellRangeAddressList(0, 0, 0, 0);
            dvHelper = sheet.GetDataValidationHelper();
            dvConstraint = dvHelper.CreateFormulaListConstraint("CHOICES");
            validation = dvHelper.CreateValidation(dvConstraint, addressList);
            sheet.AddValidationData(validation);

            addressList = new CellRangeAddressList(0, 0, 1, 1);
            dvConstraint = dvHelper.CreateFormulaListConstraint(
                    "INDIRECT(UPPER($A$1))");
            validation = dvHelper.CreateValidation(dvConstraint, addressList);
            sheet.AddValidationData(validation);

            downLoadFile_XSSF(workbook, "Excel_2007_LinkedDropDownLists.xlsx");
        }

        private void BuildDataSheet(ISheet dataSheet)
        {
            IRow row = null;
            ICell cell = null;
            IName name = null;

            // The first row will hold the data for the first validation.
            row = dataSheet.CreateRow(10);
            cell = row.CreateCell(0);
            cell.SetCellValue("Animal");
            cell = row.CreateCell(1);
            cell.SetCellValue("Vegetable");
            cell = row.CreateCell(2);
            cell.SetCellValue("Mineral");
            name = dataSheet.Workbook.CreateName();
            name.RefersToFormula = "'Linked Validations'!$A$11:$C$11";
            name.NameName = "CHOICES";

            // The next three rows will hold the data that will be used to
            // populate the second, or linked, drop down list.
            row = dataSheet.CreateRow(11);
            cell = row.CreateCell(0);
            cell.SetCellValue("Lion");
            cell = row.CreateCell(1);
            cell.SetCellValue("Tiger");
            cell = row.CreateCell(2);
            cell.SetCellValue("Leopard");
            cell = row.CreateCell(3);
            cell.SetCellValue("Elephant");
            cell = row.CreateCell(4);
            cell.SetCellValue("Eagle");
            cell = row.CreateCell(5);
            cell.SetCellValue("Horse");
            cell = row.CreateCell(6);
            cell.SetCellValue("Zebra");
            name = dataSheet.Workbook.CreateName();
            name.RefersToFormula = "'Linked Validations'!$A$12:$G$12";
            name.NameName = "ANIMAL";

            row = dataSheet.CreateRow(12);
            cell = row.CreateCell(0);
            cell.SetCellValue("Cabbage");
            cell = row.CreateCell(1);
            cell.SetCellValue("Cauliflower");
            cell = row.CreateCell(2);
            cell.SetCellValue("Potato");
            cell = row.CreateCell(3);
            cell.SetCellValue("Onion");
            cell = row.CreateCell(4);
            cell.SetCellValue("Beetroot");
            cell = row.CreateCell(5);
            cell.SetCellValue("Asparagus");
            cell = row.CreateCell(6);
            cell.SetCellValue("Spinach");
            cell = row.CreateCell(7);
            cell.SetCellValue("Chard");
            name = dataSheet.Workbook.CreateName();
            name.RefersToFormula = "'Linked Validations'!$A$13:$H$13";
            name.NameName = "VEGETABLE";

            row = dataSheet.CreateRow(13);
            cell = row.CreateCell(0);
            cell.SetCellValue("Bauxite");
            cell = row.CreateCell(1);
            cell.SetCellValue("Quartz");
            cell = row.CreateCell(2);
            cell.SetCellValue("Feldspar");
            cell = row.CreateCell(3);
            cell.SetCellValue("Shist");
            cell = row.CreateCell(4);
            cell.SetCellValue("Shale");
            cell = row.CreateCell(5);
            cell.SetCellValue("Mica");
            name = dataSheet.Workbook.CreateName();
            name.RefersToFormula = "'Linked Validations'!$A$14:$F$14";
            name.NameName = "MINERAL";
        }

        #endregion

        #region 带有每月工资报表的Excel

        protected void lnkbtnExcel_2007_MonthlySalaryReport_Click(object sender, EventArgs e)
        {
            IWorkbook workBook = new XSSFWorkbook();
            ISheet s1 = workBook.CreateSheet("Monthly Salary Report");
            IRow headerRow = s1.CreateRow(0);
            headerRow.CreateCell(0).SetCellValue("First Name");
            s1.SetColumnWidth(0, 20 * 256);
            headerRow.CreateCell(1).SetCellValue("Last Name");
            s1.SetColumnWidth(1, 20 * 256);
            headerRow.CreateCell(2).SetCellValue("Salary");
            headerRow.CreateCell(3).SetCellValue("Tax Rate");
            headerRow.CreateCell(4).SetCellValue("Tax");
            headerRow.CreateCell(5).SetCellValue("Delivery");

            int row = 1;
            GenerateRow(s1, row++, "Bill", "Zhang", 5000, 9.0 / 100);
            GenerateRow(s1, row++, "Amy", "Huang", 8000, 11.0 / 100);
            GenerateRow(s1, row++, "Tomos", "Johnson", 6000, 9.0 / 100);
            GenerateRow(s1, row++, "Macro", "Jeep", 12000, 15.0 / 100);
            s1.ForceFormulaRecalculation = true;

            downLoadFile_XSSF(workBook, "Excel_2007_MonthlySalaryReport.xlsx");
        }

        private void GenerateRow(ISheet sheet1, int rowid, string firstName, string lastName, double salaryAmount, double taxRate)
        {
            IRow row = sheet1.CreateRow(rowid);
            row.CreateCell(0).SetCellValue(firstName);  //A2
            row.CreateCell(1).SetCellValue(lastName);   //B2
            row.CreateCell(2).SetCellValue(salaryAmount);   //C2
            row.CreateCell(3).SetCellValue(taxRate);        //D2
            row.CreateCell(4).SetCellFormula(string.Format("C{0}*D{0}", rowid + 1));
            row.CreateCell(5).SetCellFormula(string.Format("C{0}-E{0}", rowid + 1));
        }

        #endregion

        //Excek 页面设置(没看出来)
        protected void lnkbtnExcel_2007_PageSetup_Click(object sender, EventArgs e)
        {
            IWorkbook wb = new XSSFWorkbook();
            ISheet sheet1 = wb.CreateSheet("new sheet");
            ISheet sheet2 = wb.CreateSheet("second sheet");

            // Set the columns to repeat from column 0 to 2 on the first sheet
            IRow row1 = sheet1.CreateRow(0);
            row1.CreateCell(0).SetCellValue(1);
            row1.CreateCell(1).SetCellValue(2);
            row1.CreateCell(2).SetCellValue(3);
            IRow row2 = sheet1.CreateRow(1);
            row2.CreateCell(1).SetCellValue(4);
            row2.CreateCell(2).SetCellValue(5);


            IRow row3 = sheet2.CreateRow(1);
            row3.CreateCell(0).SetCellValue(2.1);
            row3.CreateCell(4).SetCellValue(2.2);
            row3.CreateCell(5).SetCellValue(2.3);
            IRow row4 = sheet2.CreateRow(2);
            row4.CreateCell(4).SetCellValue(2.4);
            row4.CreateCell(5).SetCellValue(2.5);

            // Set the columns to repeat from column 0 to 2 on the first sheet
            wb.SetRepeatingRowsAndColumns(0, 0, 2, -1, -1);
            // Set the the repeating rows and columns on the second sheet.
            wb.SetRepeatingRowsAndColumns(1, 4, 5, 1, 2);

            //set the print area for the first sheet
            wb.SetPrintArea(0, 1, 2, 0, 3);

            downLoadFile_XSSF(wb, "Excel_2007_PageSetup.xlsx");
        }

        //Excel打印设置
        protected void lnkbtnExcel_2007_PrintSetup_Click(object sender, EventArgs e)
        {
            IWorkbook workBook = new XSSFWorkbook();
            ISheet sheet1 = workBook.CreateSheet("Sheet1");
            sheet1.SetMargin(MarginType.RightMargin, 0.5d);
            sheet1.SetMargin(MarginType.TopMargin, 0.6d);
            sheet1.SetMargin(MarginType.LeftMargin, 0.4d);
            sheet1.SetMargin(MarginType.BottomMargin, 0.3d);


            sheet1.PrintSetup.Copies = 3;
            sheet1.PrintSetup.NoColor = true;
            sheet1.PrintSetup.Landscape = true;
            sheet1.PrintSetup.PaperSize = (short)PaperSize.A4 + 1;

            sheet1.FitToPage = true;
            sheet1.PrintSetup.FitHeight = 2;
            sheet1.PrintSetup.FitWidth = 3;
            sheet1.IsPrintGridlines = true;

            sheet1.CreateRow(0).CreateCell(0).SetCellValue("This is a Sample");
            int x = 1;
            for (int i = 1; i <= 15; i++)
            {
                IRow row = sheet1.CreateRow(i);
                for (int j = 0; j < 15; j++)
                {
                    row.CreateCell(j).SetCellValue(x++);
                }
            }

            ISheet sheet2 = workBook.CreateSheet("Sheet2");
            sheet2.PrintSetup.Copies = 1;
            sheet2.PrintSetup.Landscape = false;
            sheet2.PrintSetup.Notes = true;
            //sheet2.PrintSetup.EndNote = true;
            //sheet2.PrintSetup.CellError = DisplayCellErrorType.ErrorAsNA;
            sheet2.PrintSetup.PaperSize = (short)PaperSize.A5 + 1;

            x = 100;
            for (int i = 1; i <= 15; i++)
            {
                IRow row = sheet2.CreateRow(i);
                for (int j = 0; j < 15; j++)
                {
                    row.CreateCell(j).SetCellValue(x++);
                }
            }

            downLoadFile_XSSF(workBook, "Excel_2007_PageSetup.xlsx");
        }

        //带有保护机制的Excel
        protected void lnkbtnExcel_2007_ProtectSheet_Click(object sender, EventArgs e)
        {
            IWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet1 = (XSSFSheet)workbook.CreateSheet("Sheet A1");

            sheet1.LockFormatRows();
            sheet1.LockFormatCells();
            sheet1.LockFormatColumns();
            sheet1.LockDeleteColumns();
            sheet1.LockDeleteRows();
            sheet1.LockInsertHyperlinks();
            sheet1.LockInsertColumns();
            sheet1.LockInsertRows();
            sheet1.ProtectSheet("password");

            downLoadFile_XSSF(workbook, "Excel_2007_ProtectSheet.xlsx");
        }

        //带有散点图的Excel
        protected void lnkbtnExcel_2007_ScatterChart_Click(object sender, EventArgs e)
        {
            IWorkbook wb = new XSSFWorkbook();
            ISheet sheet = wb.CreateSheet("Sheet 1");
            int NUM_OF_ROWS = 3;
            int NUM_OF_COLUMNS = 10;

            // Create a row and put some cells in it. Rows are 0 based.
            IRow row;
            ICell cell;
            for (int rowIndex = 0; rowIndex < NUM_OF_ROWS; rowIndex++)
            {
                row = sheet.CreateRow((short)rowIndex);
                for (int colIndex = 0; colIndex < NUM_OF_COLUMNS; colIndex++)
                {
                    cell = row.CreateCell((short)colIndex);
                    cell.SetCellValue(colIndex * (rowIndex + 1));
                }
            }

            IDrawing drawing = sheet.CreateDrawingPatriarch();
            IClientAnchor anchor = drawing.CreateAnchor(0, 0, 0, 0, 0, 5, 10, 15);

            IChart chart = drawing.CreateChart(anchor);
            NPOI.SS.UserModel.Charts.IChartLegend legend = chart.GetOrCreateLegend();
            legend.Position = (NPOI.SS.UserModel.Charts.LegendPosition.TopRight);

            NPOI.SS.UserModel.Charts.IScatterChartData<double, double> data = chart.GetChartDataFactory().CreateScatterChartData<double, double>();

            NPOI.SS.UserModel.Charts.IValueAxis bottomAxis = chart.GetChartAxisFactory().CreateValueAxis(NPOI.SS.UserModel.Charts.AxisPosition.Bottom);
            NPOI.SS.UserModel.Charts.IValueAxis leftAxis = chart.GetChartAxisFactory().CreateValueAxis(NPOI.SS.UserModel.Charts.AxisPosition.Left);
            leftAxis.SetCrosses(NPOI.SS.UserModel.Charts.AxisCrosses.AutoZero);

            NPOI.SS.UserModel.Charts.IChartDataSource<double> xs = NPOI.SS.UserModel.Charts.DataSources.FromNumericCellRange(sheet, new CellRangeAddress(0, 0, 0, NUM_OF_COLUMNS - 1));
            NPOI.SS.UserModel.Charts.IChartDataSource<double> ys1 = NPOI.SS.UserModel.Charts.DataSources.FromNumericCellRange(sheet, new CellRangeAddress(1, 1, 0, NUM_OF_COLUMNS - 1));
            NPOI.SS.UserModel.Charts.IChartDataSource<double> ys2 = NPOI.SS.UserModel.Charts.DataSources.FromNumericCellRange(sheet, new CellRangeAddress(2, 2, 0, NUM_OF_COLUMNS - 1));


            var s1 = data.AddSerie(xs, ys1);
            s1.SetTitle("title1");
            var s2 = data.AddSerie(xs, ys2);
            s2.SetTitle("title2");
            chart.Plot(data, bottomAxis, leftAxis);

            downLoadFile_XSSF(wb, "Excel_2007_ScatterChart.xlsx");
        }

        //设置单元格的长和宽
        protected void lnkbtnExcel_2007_SetWidthAndHeight_Click(object sender, EventArgs e)
        {
            IWorkbook workbook = new XSSFWorkbook();
            ISheet sheet1 = workbook.CreateSheet("Sheet1");
            //set the width of columns
            sheet1.SetColumnWidth(0, 50 * 256);
            sheet1.SetColumnWidth(1, 100 * 256);
            sheet1.SetColumnWidth(2, 150 * 256);

            //set the width of height
            sheet1.CreateRow(0).Height = 100 * 20;
            sheet1.CreateRow(1).Height = 200 * 20;
            sheet1.CreateRow(2).Height = 300 * 20;

            downLoadFile_XSSF(workbook, "Excel_2007_SetWidthAndHeight.xlsx");
        }

        //带有拆分和冻结窗口的Excel
        protected void lnkbtnExcel_2007_SplitAndFreezePanes_Click(object sender, EventArgs e)
        {
            IWorkbook workbook = new XSSFWorkbook();
            ISheet sheet1 = workbook.CreateSheet("new sheet");
            ISheet sheet2 = workbook.CreateSheet("second sheet");
            ISheet sheet3 = workbook.CreateSheet("third sheet");
            ISheet sheet4 = workbook.CreateSheet("fourth sheet");

            // Freeze just one row
            sheet1.CreateFreezePane(0, 1, 0, 1);
            // Freeze just one column
            sheet2.CreateFreezePane(1, 0, 1, 0);
            // Freeze the columns and rows (forget about scrolling position of the lower right quadrant).
            sheet3.CreateFreezePane(2, 2);
            // Create a split with the lower left side being the active quadrant
            sheet4.CreateSplitPane(2000, 2000, 0, 0, PanePosition.LowerLeft);

            downLoadFile_XSSF(workbook, "Excel_2007_SplitAndFreezePanes.xlsx");
        }

        //写性能测试(7W行数据)
        protected void lnkbtnExcel_2007_WritePerformanceTest_Click(object sender, EventArgs e)
        {
            IWorkbook workbook = new XSSFWorkbook();
            ISheet sheet1 = workbook.CreateSheet("Sheet1");
            sheet1.CreateRow(0).CreateCell(0).SetCellValue("This is a Sample");
            int x = 1;

            Console.WriteLine("Start at " + DateTime.Now.ToString());
            for (int i = 1; i <= 70000; i++)
            {
                IRow row = sheet1.CreateRow(i);
                for (int j = 0; j < 15; j++)
                {
                    row.CreateCell(j).SetCellValue(x++);
                }
            }

            downLoadFile_XSSF(workbook, "Excel_2007_WritePerformanceTest.xlsx");
        }

        #endregion

        #region Word 2007

        //创建一个空的Word文档
        protected void lnkbtnWord_2007_CreateEmptyDocument_Click(object sender, EventArgs e)
        {
            XWPFDocument doc = new XWPFDocument();
            doc.CreateParagraph();

            downLoadFile_XWPF(doc, "Word_2007_CreateEmptyDocument.docx");
        }

        //创建一个带有图片的Word文档
        protected void lnkbtnWord_2007_InsertPicturesInWord_Click(object sender, EventArgs e)
        {
            XWPFDocument doc = new XWPFDocument();
            XWPFParagraph p2 = doc.CreateParagraph();
            XWPFRun r2 = p2.CreateRun();
            r2.SetText("test");

            var img = System.Drawing.Image.FromFile(Server.MapPath("Images/HumpbackWhale.jpg"));

            var widthEmus = (int)(400.0 * 9525);
            var heightEmus = (int)(300.0 * 9525);

            using (FileStream picData = new FileStream(Server.MapPath("Images/HumpbackWhale.jpg"), FileMode.Open, FileAccess.Read))
            {
                r2.AddPicture(picData, (int)NPOI.XWPF.UserModel.PictureType.PNG, "image1", widthEmus, heightEmus);
            }

            downLoadFile_XWPF(doc, "Word_2007_InsertPicturesInWord.docx");
        }

        //创建一个简单的Word文档
        protected void lnkbtnWord_2007_SimpleDocument_Click(object sender, EventArgs e)
        {
            XWPFDocument doc = new XWPFDocument();

            XWPFParagraph p1 = doc.CreateParagraph();
            p1.Alignment = ParagraphAlignment.CENTER;
            p1.BorderBottom = Borders.DOUBLE;
            p1.BorderTop = Borders.DOUBLE;

            p1.BorderRight = Borders.DOUBLE;
            p1.BorderLeft = Borders.DOUBLE;
            p1.BorderBetween = Borders.SINGLE;

            p1.VerticalAlignment = TextAlignment.TOP;

            XWPFRun r1 = p1.CreateRun();
            r1.SetBold(true);
            r1.SetText("The quick brown fox");
            r1.SetBold(true);
            r1.FontFamily = "Courier";
            r1.SetUnderline(UnderlinePatterns.DotDotDash);
            r1.SetTextPosition(100);

            XWPFParagraph p2 = doc.CreateParagraph();
            p2.Alignment = ParagraphAlignment.RIGHT;

            //BORDERS
            p2.BorderBottom = Borders.DOUBLE;
            p2.BorderTop = Borders.DOUBLE;
            p2.BorderRight = Borders.DOUBLE;
            p2.BorderLeft = Borders.DOUBLE;
            p2.BorderBetween = Borders.SINGLE;

            XWPFRun r2 = p2.CreateRun();
            r2.SetText("jumped over the lazy dog");
            r2.SetStrike(true);
            r2.FontSize = 20;

            XWPFRun r3 = p2.CreateRun();
            r3.SetText("and went away");
            r3.SetStrike(true);
            r3.FontSize = 20;
            r3.Subscript = NPOI.XWPF.UserModel.VerticalAlign.SUPERSCRIPT;


            XWPFParagraph p3 = doc.CreateParagraph();
            p3.IsWordWrap = true;
            p3.IsPageBreak = true;

            //p3.SetAlignment(ParagraphAlignment.DISTRIBUTE);
            p3.Alignment = ParagraphAlignment.BOTH;
            p3.SpacingLineRule = LineSpacingRule.EXACT;

            p3.IndentationFirstLine = 600;


            XWPFRun r4 = p3.CreateRun();
            r4.SetTextPosition(20);
            r4.SetText("To be, or not to be: that is the question: "
                    + "Whether 'tis nobler in the mind to suffer "
                    + "The slings and arrows of outrageous fortune, "
                    + "Or to take arms against a sea of troubles, "
                    + "And by opposing end them? To die: to sleep; ");
            r4.AddBreak(BreakType.PAGE);
            r4.SetText("No more; and by a sleep to say we end "
                    + "The heart-ache and the thousand natural shocks "
                    + "That flesh is heir to, 'tis a consummation "
                    + "Devoutly to be wish'd. To die, to sleep; "
                    + "To sleep: perchance to dream: ay, there's the rub; "
                    + ".......");
            r4.IsItalic = true;
            //This would imply that this break shall be treated as a simple line break, and break the line after that word:

            XWPFRun r5 = p3.CreateRun();
            r5.SetTextPosition(-10);
            r5.SetText("For in that sleep of death what dreams may come");
            r5.AddCarriageReturn();
            r5.SetText("When we have shuffled off this mortal coil,"
                    + "Must give us pause: there's the respect"
                    + "That makes calamity of so long life;");
            r5.AddBreak();
            r5.SetText("For who would bear the whips and scorns of time,"
                    + "The oppressor's wrong, the proud man's contumely,");

            r5.AddBreak(BreakClear.ALL);
            r5.SetText("The pangs of despised love, the law's delay,"
                    + "The insolence of office and the spurns" + ".......");

            downLoadFile_XWPF(doc, "Word_2007_SimpleDocument.docx");
        }

        //带有表格的Word文档
        protected void lnkbtnWord_2007_SimpleTable_Click(object sender, EventArgs e)
        {
            XWPFDocument doc = new XWPFDocument();

            XWPFTable table = doc.CreateTable(3, 3);

            table.GetRow(1).GetCell(1).SetText("EXAMPLE OF TABLE");


            XWPFParagraph p1 = doc.CreateParagraph();

            XWPFRun r1 = p1.CreateRun();
            r1.SetBold(true);
            r1.SetText("The quick brown fox");
            r1.SetBold(true);
            r1.FontFamily = "Courier";
            r1.SetUnderline(UnderlinePatterns.DotDotDash);
            r1.SetTextPosition(100);

            table.GetRow(0).GetCell(0).SetParagraph(p1);


            table.GetRow(2).GetCell(2).SetText("only text");

            downLoadFile_XWPF(doc, "Word_2007_SimpleTable.docx");
        }

        //更新Word
        protected void lnkbtnWord_2007_UpdateEmbeddedDoc_Click(object sender, EventArgs e)
        {
            string path = Server.MapPath("WordMode/WordMode.docx");
            using (FileStream file = new FileStream(path, FileMode.Open))
            {
                XWPFDocument doc = new XWPFDocument(file);

                IWorkbook workbook = null;
                ISheet sheet = null;
                IRow row = null;
                ICell cell = null;
                PackagePart pPart = null;
                IEnumerator<PackagePart> pIter = null;
                List<PackagePart> embeddedDocs = doc.GetAllEmbedds();
                if (embeddedDocs != null && embeddedDocs.Count != 0)
                {
                    pIter = embeddedDocs.GetEnumerator();
                    while (pIter.MoveNext())
                    {
                        pPart = pIter.Current;

                        workbook = WorkbookFactory.Create(pPart.GetInputStream());
                        sheet = workbook.GetSheetAt(0);
                        row = sheet.GetRow(0);
                        cell = row.GetCell(0);
                        cell.SetCellValue("xxxxxxxxxxxxxxxxxxxxx");
                        workbook.Write(pPart.GetOutputStream());
                    }
                    downLoadFile_XWPF(doc, "Word_2007_UpdateEmbeddedDoc.docx");
                }
            }
        }

        //带有序列的Word文档
        protected void lnkbtnWord_2007_CreateBullet_Click(object sender, EventArgs e)
        {
            XWPFDocument doc = new XWPFDocument();
            //simple bullet
            XWPFNumbering numbering = doc.CreateNumbering();

            string abstractNumId = numbering.AddAbstractNum();
            string numId = numbering.AddNum(abstractNumId);

            XWPFParagraph p0 = doc.CreateParagraph();
            XWPFRun r0 = p0.CreateRun();
            r0.SetText("simple bullet");
            r0.SetBold(true);
            r0.FontFamily = "Courier";
            r0.FontSize = 12;

            XWPFParagraph p1 = doc.CreateParagraph();
            XWPFRun r1 = p1.CreateRun();
            r1.SetText("first, create paragraph and run, set text");
            p1.SetNumID(numId);

            XWPFParagraph p2 = doc.CreateParagraph();
            XWPFRun r2 = p2.CreateRun();
            r2.SetText("second, call XWPFDocument.CreateNumbering() to create numbering");
            p2.SetNumID(numId);

            XWPFParagraph p3 = doc.CreateParagraph();
            XWPFRun r3 = p3.CreateRun();
            r3.SetText("third, add AbstractNum[numbering.AddAbstractNum()] and Num(numbering.AddNum(abstractNumId))");
            p3.SetNumID(numId);

            XWPFParagraph p4 = doc.CreateParagraph();
            XWPFRun r4 = p4.CreateRun();
            r4.SetText("next, call XWPFParagraph.SetNumID(numId) to set paragraph property, CT_P.pPr.numPr");
            p4.SetNumID(numId);

            //multi level
            abstractNumId = numbering.AddAbstractNum();
            numId = numbering.AddNum(abstractNumId);
            doc.CreateParagraph();
            doc.CreateParagraph();

            p1 = doc.CreateParagraph();
            r1 = p1.CreateRun();
            r1.SetText("multi level bullet");
            r1.SetBold(true);
            r1.FontFamily = "Courier";
            r1.FontSize = 12;

            p1 = doc.CreateParagraph();
            r1 = p1.CreateRun();
            r1.SetText("first");
            p1.SetNumID(numId, "0");
            p1 = doc.CreateParagraph();
            r1 = p1.CreateRun();
            r1.SetText("first-first");
            p1.SetNumID(numId, "1");
            p1 = doc.CreateParagraph();
            r1 = p1.CreateRun();
            r1.SetText("first-second");
            p1.SetNumID(numId, "1");
            p1 = doc.CreateParagraph();
            r1 = p1.CreateRun();
            r1.SetText("first-third");
            p1.SetNumID(numId, "1");
            p1 = doc.CreateParagraph();
            r1 = p1.CreateRun();
            r1.SetText("second");
            p1.SetNumID(numId, "0");
            p1 = doc.CreateParagraph();
            r1 = p1.CreateRun();
            r1.SetText("second-first");
            p1.SetNumID(numId, "1");
            p1 = doc.CreateParagraph();
            r1 = p1.CreateRun();
            r1.SetText("second-second");
            p1.SetNumID(numId, "1");
            p1 = doc.CreateParagraph();
            r1 = p1.CreateRun();
            r1.SetText("second-third");
            p1.SetNumID(numId, "1");
            p1 = doc.CreateParagraph();
            r1 = p1.CreateRun();
            r1.SetText("second-third-first");
            p1.SetNumID(numId, "2");
            p1 = doc.CreateParagraph();
            r1 = p1.CreateRun();
            r1.SetText("second-third-second");
            p1.SetNumID(numId, "2");
            p1 = doc.CreateParagraph();
            r1 = p1.CreateRun();
            r1.SetText("third");
            p1.SetNumID(numId, "0");

            downLoadFile_XWPF(doc, "Word_2007_CreateBullet.docx");
        }

        //带有合并单元格的Word文档
        protected void lnkbtnWord_2007_CreateTable_Click(object sender, EventArgs e)
        {
            XWPFDocument doc = new XWPFDocument();
            XWPFParagraph p0 = doc.CreateParagraph();
            XWPFRun r0 = p0.CreateRun();
            r0.SetText("DOCX表");

            XWPFTable table = doc.CreateTable(1, 3);//创建一行3列表
            table.GetRow(0).GetCell(0).SetText("111");
            table.GetRow(0).GetCell(1).SetText("222");
            table.GetRow(0).GetCell(2).SetText("333");

            XWPFTableRow m_Row = table.CreateRow();//创建一行
            m_Row = table.CreateRow();//创建一行
            m_Row.GetCell(0).SetText("211");

            //合并单元格
            m_Row = table.InsertNewTableRow(0);//表头插入一行
            XWPFTableCell cell = m_Row.CreateCell();//创建一个单元格,创建单元格时就创建了一个CT_P
            CT_Tc cttc = cell.GetCTTc();
            CT_TcPr ctPr = cttc.AddNewTcPr();
            ctPr.gridSpan.val = "3";//合并3列
            cttc.GetPList()[0].AddNewPPr().AddNewJc().val = ST_Jc.center;
            cttc.GetPList()[0].AddNewR().AddNewT().Value = "abc";

            XWPFTableRow td3 = table.InsertNewTableRow(table.Rows.Count - 1);//插入行
            cell = td3.CreateCell();
            cttc = cell.GetCTTc();
            ctPr = cttc.AddNewTcPr();
            ctPr.gridSpan.val = "3";
            cttc.GetPList()[0].AddNewPPr().AddNewJc().val = ST_Jc.center;
            cttc.GetPList()[0].AddNewR().AddNewT().Value = "qqq";

            //表增加行，合并列
            CT_Row m_NewRow = new CT_Row();
            m_Row = new XWPFTableRow(m_NewRow, table);
            table.AddRow(m_Row); //必须要！！！
            cell = m_Row.CreateCell();
            cttc = cell.GetCTTc();
            ctPr = cttc.AddNewTcPr();
            ctPr.gridSpan.val = "3";
            cttc.GetPList()[0].AddNewPPr().AddNewJc().val = ST_Jc.center;
            cttc.GetPList()[0].AddNewR().AddNewT().Value = "sss";

            //表未增加行，合并2列，合并2行
            //1行
            m_NewRow = new CT_Row();
            m_Row = new XWPFTableRow(m_NewRow, table);
            table.AddRow(m_Row);
            cell = m_Row.CreateCell();
            cttc = cell.GetCTTc();
            ctPr = cttc.AddNewTcPr();
            ctPr.gridSpan.val = "2";
            ctPr.AddNewVMerge().val = ST_Merge.restart;//合并行
            ctPr.AddNewVAlign().val = ST_VerticalJc.center;//垂直居中
            cttc.GetPList()[0].AddNewPPr().AddNewJc().val = ST_Jc.center;
            cttc.GetPList()[0].AddNewR().AddNewT().Value = "xxx";
            cell = m_Row.CreateCell();
            cell.SetText("ddd");
            //2行，多行合并类似
            m_NewRow = new CT_Row();
            m_Row = new XWPFTableRow(m_NewRow, table);
            table.AddRow(m_Row);
            cell = m_Row.CreateCell();
            cttc = cell.GetCTTc();
            ctPr = cttc.AddNewTcPr();
            ctPr.gridSpan.val = "2";
            ctPr.AddNewVMerge().val = ST_Merge.@continue;//合并行
            cell = m_Row.CreateCell();
            cell.SetText("kkk");
            ////3行
            //m_NewRow = new CT_Row();
            //m_Row = new XWPFTableRow(m_NewRow, table);
            //table.AddRow(m_Row);
            //cell = m_Row.CreateCell();
            //cttc = cell.GetCTTc();
            //ctPr = cttc.AddNewTcPr();
            //ctPr.gridSpan.val = "2";
            //ctPr.AddNewVMerge().val = ST_Merge.@continue;
            //cell = m_Row.CreateCell();
            //cell.SetText("hhh");

            downLoadFile_XWPF(doc, "Word_2007_CreateTable.docx");
        }

        #endregion
    }
}

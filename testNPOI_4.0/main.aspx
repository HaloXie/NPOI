<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="main.aspx.cs" Inherits="testNPOI_4._0.main" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title></title>
    <link href="CSS/buttons.css" rel="stylesheet" />
    <link href="CSS/page.css" rel="stylesheet" />
    <link href="CSS/table.css" rel="stylesheet" />
    <style type="text/css">
        .auto-style1
        {
            color: #FF0000;
            padding-left: 20px;
            float: left;
        }

        .auto-style2
        {
            padding-left: 40px;
            float: left;
            padding-bottom: 20px;
        }
    </style>
</head>
<body>
    <form id="form1" runat="server">
        <div style="margin: 0 auto; margin-top: 50px; width: 1000px">
            <div style="text-align: center">
                <b style="color: #0000FF">NPOI 4.0导出程序：由于部分的Word在2.0中实现不了，现在改为4.0中尝试一下</b><br />
                <div id="btns">
                    <div id="buttonContainer">
                        <a href="#Excel_2003" class="button big green">&nbsp;&nbsp;Excel 2003:HSSF&nbsp;&nbsp;</a>
                        <a href="#Excel_2007" class="button big green">&nbsp;&nbsp;Excel 2007&nbsp;&nbsp;</a>
                        <a href="#Word_2007" class="button big orange">&nbsp;&nbsp;Word 2007&nbsp;&nbsp;</a>
                        <a href="#XML" class="button big orange">&nbsp;&nbsp;压缩文件:OOXML&nbsp;&nbsp;</a>
                        <a href="#POIFS" class="button big green">&nbsp;&nbsp;Office 自定义属性:POIFS&nbsp;&nbsp;</a>
                        <a href="#NPOI_1" class="button big green">&nbsp;&nbsp;NPOI 1.2.*说明&nbsp;&nbsp;</a>
                        <a href="#NPOI_2" class="button big orange">&nbsp;&nbsp;NPOI 2.*说明&nbsp;&nbsp;</a>
                        <div style="display: none">
                            <br />
                            <a href="#" class="button big blue">Big Button</a>
                            <a href="#" class="button big green">Big Button</a>
                            <a href="#" class="button big orange">Big Button</a>
                            <a href="#" class="button big gray">Big Button</a>
                            <br />
                            <a href="#" class="button blue medium">Medium Button</a>
                            <a href="#" class="button green medium">Medium Button</a>
                            <a href="#" class="button orange medium">Medium Button</a>
                            <a href="#" class="button gray medium">Medium Button</a>
                            <br />
                            <a href="#" class="button small blue">Small Button</a>
                            <a href="#" class="button small green">Small Button</a>
                            <a href="#" class="button small blue rounded">Rounded</a>
                            <a href="#" class="button small orange">Small Button</a>
                            <a href="#" class="button small gray">Small Button</a>
                            <a href="#" class="button small green rounded">Rounded</a>
                        </div>
                    </div>
                </div>
            </div>
            <br />
            <br />
            <div id="Excel_2003">
                <b style="color: #0000FF">关于Excel 2003导出：HSSF</b>
                <div style="border: 1px gray dashed; margin-top: 10px; height: auto; overflow: hidden">
                    <div style="padding-top: 30px; padding-bottom: 20px;">
                        <span class="auto-style1">*主要是针对Excel 2003的一些基本操作</span><br />
                        <br />
                        <div class="auto-style2">
                            <ul>
                                <li>
                                    <asp:LinkButton ID="lnkbtnExcel_2003_EmptyExcel" runat="server" ForeColor="Blue" OnClick="lnkbtnExcel_2003_EmptyExcel_Click">创建一个空白的Excel</asp:LinkButton></li>
                                <li>
                                    <asp:LinkButton ID="lnkbtnExcel_2003_HyperLink" runat="server" ForeColor="Blue" OnClick="lnkbtnExcel_2003_HyperLink_Click">带有HyperLink的Excel</asp:LinkButton></li>
                                <li>
                                    <asp:LinkButton ID="lnkbtnExcel_2003_Font" runat="server" ForeColor="Blue" OnClick="lnkbtnExcel_2003_Font_Click">带有字体应用的Excel</asp:LinkButton></li>
                                <li>
                                    <asp:LinkButton ID="lnkbtnExcel_2003_AutoWidth" runat="server" ForeColor="Blue" OnClick="lnkbtnExcel_2003_AutoWidth_Click">带有自动适应列宽的Excel</asp:LinkButton></li>
                                <li>
                                    <asp:LinkButton ID="lnkbtnExcel_2003_BusinessPlan" runat="server" ForeColor="Blue" OnClick="lnkbtnExcel_2003_BusinessPlan_Click">包含业务计划的Excel</asp:LinkButton>
                                </li>
                                <li>
                                    <asp:LinkButton ID="lnkbtnExcel_2003_Calendar" runat="server" ForeColor="Blue" OnClick="lnkbtnExcel_2003_Calendar_Click">包含日历的Excel</asp:LinkButton>
                                </li>
                                <li>
                                    <asp:LinkButton ID="lnkbtnExcel_2003_SheetColor" runat="server" ForeColor="Blue" OnClick="lnkbtnExcel_2003_SheetColor_Click">可以改变sheet颜色的Excel</asp:LinkButton></li>
                                <li>
                                    <asp:LinkButton ID="lnkbtnExcel_2003_ColorMatrixTable" runat="server" ForeColor="Blue" OnClick="lnkbtnExcel_2003_ColorMatrixTable_Click">带有彩色矩阵的Excel</asp:LinkButton></li>
                                <li>
                                    <asp:LinkButton ID="lnkbtnExcel_2003_ConditionalFormat" runat="server" ForeColor="Blue" OnClick="lnkbtnExcel_2003_ConditionalFormat_Click">带有规则格式的Excel</asp:LinkButton></li>
                                <li>
                                    <asp:LinkButton ID="lnkbtnExcel_2003_ExcelToHtml" runat="server" ForeColor="Blue" OnClick="lnkbtnExcel_2003_ExcelToHtml_Click">把Excel转换成Html</asp:LinkButton></li>
                                <li>
                                    <asp:LinkButton ID="lnkbtnExcel_2003_CopyRowsAndCell" runat="server" ForeColor="Blue" OnClick="lnkbtnExcel_2003_CopyRowsAndCell_Click">复制别的Excel的行和列产生新的Excel</asp:LinkButton></li>
                                <li>
                                    <asp:LinkButton ID="lnkbtnExcel_2003_CopySheet" runat="server" ForeColor="Blue" OnClick="lnkbtnExcel_2003_CopySheet_Click">复制一个Excel的Sheet产生新的Excel(用于把多个Excel合成一个Excel)</asp:LinkButton></li>
                                <li>
                                    <asp:LinkButton ID="lnkbtnExcel_2003_DropDownList" runat="server" ForeColor="Blue" OnClick="lnkbtnExcel_2003_DropDownList_Click">带有DropDownList的Excel</asp:LinkButton></li>
                                <li>
                                    <asp:LinkButton ID="lnkbtnExcel_2003_HeadAndFoot" runat="server" ForeColor="Blue" OnClick="lnkbtnExcel_2003_HeadAndFoot_Click">带有页眉和页脚的Excel</asp:LinkButton></li>
                                <li>
                                    <asp:LinkButton ID="lnkbtnExcel_2003_CustomColor" runat="server" ForeColor="Blue" OnClick="lnkbtnExcel_2003_CustomColor_Click">带有自定义颜色的Excel</asp:LinkButton></li>
                                <li>
                                    <asp:LinkButton ID="lnkbtnExcel_2003_ShowGridLines" runat="server" ForeColor="Blue" OnClick="lnkbtnExcel_2003_ShowGridLines_Click">带有网格线的Excel</asp:LinkButton></li>
                                <li>
                                    <asp:LinkButton ID="lnkbtnExcel_2003_Drawing" runat="server" ForeColor="Blue" OnClick="lnkbtnExcel_2003_Drawing_Click">带有绘画的Excel</asp:LinkButton></li>
                                <li>
                                    <asp:LinkButton ID="lnkbtnExcel_2003_AutoFilter" runat="server" ForeColor="Blue" OnClick="lnkbtnExcel_2003_AutoFilter_Click">带有排序筛选的Excel</asp:LinkButton></li>
                                <li>
                                    <asp:LinkButton ID="lnkbtnExcel_2003_ExtractPictures" runat="server" ForeColor="Blue" OnClick="lnkbtnExcel_2003_ExtractPictures_Click">提取Excel中的图片</asp:LinkButton></li>
                                <li>
                                    <asp:LinkButton ID="lnkbtnExcel_2003_ExtractString" runat="server" ForeColor="Blue" OnClick="lnkbtnExcel_2003_ExtractString_Click">提取Excel中的字符串</asp:LinkButton></li>
                                <li>
                                    <asp:LinkButton ID="lnkbtnExcel_2003_FillBackground" runat="server" ForeColor="Blue" OnClick="lnkbtnExcel_2003_FillBackground_Click">带有背景填充的Excel</asp:LinkButton></li>
                                <li>
                                    <asp:LinkButton ID="lnkbtnExcel_2003_GenerateFromTemplate" runat="server" ForeColor="Blue" OnClick="lnkbtnExcel_2003_GenerateFromTemplate_Click">读取模板生成Excel</asp:LinkButton></li>
                                <li>
                                    <asp:LinkButton ID="lnkbtnExcel_2003_FromTemplateChart" runat="server" ForeColor="Blue" OnClick="lnkbtnExcel_2003_FromTemplateChart_Click">读取模板生成Excel_Chart</asp:LinkButton></li>
                                <li>
                                    <asp:LinkButton ID="lnkbtnExcel_2003_GroupRowAndColumn" runat="server" ForeColor="Blue" OnClick="lnkbtnExcel_2003_GroupRowAndColumn_Click">带有组(行,列)的Excel</asp:LinkButton></li>
                                <li>
                                    <asp:LinkButton ID="lnkbtnExcel_2003_HideRowAndColumn" runat="server" ForeColor="Blue" OnClick="lnkbtnExcel_2003_HideRowAndColumn_Click">带有隐藏行和列的Excel</asp:LinkButton></li>
                                <li>
                                    <asp:LinkButton ID="lnkbtnExcel_2003_ImportExcel" runat="server" ForeColor="Blue" OnClick="lnkbtnExcel_2003_ImportExcel_Click">导入Excel数据,转换成dataTable在读取导出</asp:LinkButton></li>
                                <li>
                                    <asp:LinkButton ID="lnkbtnExcel_2003_Image" runat="server" ForeColor="Blue" OnClick="lnkbtnExcel_2003_Image_Click">带有图片的Excel</asp:LinkButton></li>
                                <li>
                                    <asp:LinkButton ID="lnkbtnExcel_2003_LoanCalculator" runat="server" ForeColor="Blue" OnClick="lnkbtnExcel_2003_LoanCalculator_Click">带有贷款计算器的Excel</asp:LinkButton></li>
                                <li>
                                    <asp:LinkButton ID="lnkbtnExcel_2003_MergeCells" runat="server" ForeColor="Blue" OnClick="lnkbtnExcel_2003_MergeCells_Click">带有合并单元格的Excel</asp:LinkButton></li>
                                <li>
                                    <asp:LinkButton ID="lnkbtnExcel_2003_Plication" runat="server" ForeColor="Blue" OnClick="lnkbtnExcel_2003_Plication_Click">带有折叠效果的Excel(导出的时候，没有折叠效果，倒是有公式)</asp:LinkButton></li>
                                <li>
                                    <asp:LinkButton ID="lnkbtnExcel_2003_NumberFormat" runat="server" ForeColor="Blue" OnClick="lnkbtnExcel_2003_NumberFormat_Click">带有数字格式的Excel</asp:LinkButton></li>
                                <li>
                                    <asp:LinkButton ID="lnkbtnExcel_2003_ProtectSheet" runat="server" ForeColor="Blue" OnClick="lnkbtnExcel_2003_ProtectSheet_Click">带有保护机制的Excel</asp:LinkButton></li>
                                <li>
                                    <asp:LinkButton ID="lnkbtnExcel_2003_RepeatingRowsAndColumns" runat="server" ForeColor="Blue" OnClick="lnkbtnExcel_2003_RepeatingRowsAndColumns_Click">带有重复行和列的Excel</asp:LinkButton></li>
                                <li>
                                    <asp:LinkButton ID="lnkbtnExcel_2003_RotateText" runat="server" ForeColor="Blue" OnClick="lnkbtnExcel_2003_RotateText_Click">带有文字转换的Excel</asp:LinkButton></li>
                                <li>
                                    <asp:LinkButton ID="lnkbtnExcel_2003_SetActiveCellRange" runat="server" ForeColor="Blue" OnClick="lnkbtnExcel_2003_SetActiveCellRange_Click">带有设置单元格活动范围的Excel</asp:LinkButton></li>
                                <li>
                                    <asp:LinkButton ID="lnkbtnExcel_2003_SetAlignment" runat="server" ForeColor="Blue" OnClick="lnkbtnExcel_2003_SetAlignment_Click">设置对齐的Excel</asp:LinkButton></li>
                                <li>
                                    <asp:LinkButton ID="lnkbtnExcel_2003_SetBordersOfRegion" runat="server" ForeColor="Blue" OnClick="lnkbtnExcel_2003_SetBordersOfRegion_Click">设置边框区域的Excel</asp:LinkButton></li>
                                <li>
                                    <asp:LinkButton ID="lnkbtnExcel_2003_SetBorderStyle" runat="server" ForeColor="Blue" OnClick="lnkbtnExcel_2003_SetBorderStyle_Click">设置边框样式的Excel</asp:LinkButton></li>
                                <li>
                                    <asp:LinkButton ID="lnkbtnExcel_2003_SetCellComment" runat="server" ForeColor="Blue" OnClick="lnkbtnExcel_2003_SetCellComment_Click">设置单元格的评论/批注的Excel</asp:LinkButton></li>
                                <li>
                                    <asp:LinkButton ID="lnkbtnExcel_2003_SetCellValues" runat="server" ForeColor="Blue" OnClick="lnkbtnExcel_2003_SetCellValues_Click">设置单元格的值的Excel</asp:LinkButton></li>
                                <li>
                                    <asp:LinkButton ID="lnkbtnExcel_2003_SetDateCell" runat="server" ForeColor="Blue" OnClick="lnkbtnExcel_2003_SetDateCell_Click">设置日期格式的单元格的Excel</asp:LinkButton></li>
                                <li>
                                    <asp:LinkButton ID="lnkbtnExcel_2003_SetPrintArea" runat="server" ForeColor="Blue" OnClick="lnkbtnExcel_2003_SetPrintArea_Click">设置打印区域的Excel</asp:LinkButton></li>
                                <li>
                                    <asp:LinkButton ID="lnkbtnExcel_2003_SetPrintSettings" runat="server" ForeColor="Blue" OnClick="lnkbtnExcel_2003_SetPrintSettings_Click">设置打印设置的Excel</asp:LinkButton></li>
                                <li>
                                    <asp:LinkButton ID="lnkbtnExcel_2003_SetWidthAndHeight" runat="server" ForeColor="Blue" OnClick="lnkbtnExcel_2003_SetWidthAndHeight_Click">设置宽度和高度的Excel</asp:LinkButton></li>
                                <li>
                                    <asp:LinkButton ID="lnkbtnExcel_2003_ShrinkToFitColumn" runat="server" ForeColor="Blue" OnClick="lnkbtnExcel_2003_ShrinkToFitColumn_Click">缩小到合适列的Excel</asp:LinkButton></li>
                                <li>
                                    <asp:LinkButton ID="lnkbtnExcel_2003_SplitAndFreezePanes" runat="server" ForeColor="Blue" OnClick="lnkbtnExcel_2003_SplitAndFreezePanes_Click">带有拆分和冻结窗格的Excel</asp:LinkButton></li>
                                <li>
                                    <asp:LinkButton ID="lnkbtnExcel_2003_TimeSheetDemo" runat="server" ForeColor="Blue" OnClick="lnkbtnExcel_2003_TimeSheetDemo_Click">带有时间片演示的Excel</asp:LinkButton></li>
                                <li>
                                    <asp:LinkButton ID="lnkbtnExcel_2003_UseBasicFormula" runat="server" ForeColor="Blue" OnClick="lnkbtnExcel_2003_UseBasicFormula_Click">带有基本公式的Excel</asp:LinkButton></li>
                                <li>
                                    <asp:LinkButton ID="lnkbtnExcel_2003_UseNewLinesInCells" runat="server" ForeColor="Blue" OnClick="lnkbtnExcel_2003_UseNewLinesInCells_Click">在单元格中使用换行的Excel</asp:LinkButton></li>
                                <li>
                                    <asp:LinkButton ID="lnkbtnExcel_2003_ZoomSheet" runat="server" ForeColor="Blue" OnClick="lnkbtnExcel_2003_ZoomSheet_Click">带有放大缩小的Sheet的Excel</asp:LinkButton></li>
                                <li>
                                    <asp:LinkButton ID="lnkbtnExcel_2003_SetPassword" runat="server" ForeColor="Blue" OnClick="lnkbtnExcel_2003_SetPassword_Click">带有可选择性设置密码的Excel</asp:LinkButton></li>
                            </ul>
                        </div>
                    </div>
                </div>
            </div>
            <div id="Excel_2007" style="margin-top: 60px;">
                <b style="color: #0000FF">关于Excel导出 2007</b>
                <div style="width: 100%; height: auto; border: 1px gray dashed; margin-top: 10px; overflow: hidden">
                    <div style="padding-top: 30px; padding-bottom: 20px;">
                        <span class="auto-style1">*主要是针对Excel 2007的一些基本操作</span><br />
                        <br />
                        <div class="auto-style2">
                            <ul>
                                <li>
                                    <asp:LinkButton ID="lnkbtnExcel_2007_HyperLink" runat="server" ForeColor="Blue" OnClick="lnkbtnExcel_2007_HyperLink_Click">带有HyperLink的Excel</asp:LinkButton></li>
                                <li>
                                    <asp:LinkButton ID="lnkbtnExcel_2007_Font" runat="server" ForeColor="Blue" OnClick="lnkbtnExcel_2007_Font_Click">带有字体应用的Excel</asp:LinkButton></li>
                                <li>
                                    <asp:LinkButton ID="lnkbtnExcel_2007_ApplyTable" runat="server" ForeColor="Blue" OnClick="lnkbtnExcel_2007_ApplyTable_Click">带有表格的Excel(失败！！！)</asp:LinkButton></li>
                                <li>
                                    <asp:LinkButton ID="lbkbtnExcel_2007_BigGrid" runat="server" ForeColor="Blue" OnClick="lbkbtnExcel_2007_BigGrid_Click">带有1W行数据Excel导出实验</asp:LinkButton>
                                </li>
                                <li>
                                    <asp:LinkButton ID="lnkbtnExcel_2007_BorderStyles" runat="server" ForeColor="Blue" OnClick="lnkbtnExcel_2007_BorderStyles_Click">带有边框样式的Excel</asp:LinkButton>
                                </li>
                                <li>
                                    <asp:LinkButton ID="lnkbtnExcel_2007_ColorfulMatrix" runat="server" ForeColor="Blue" OnClick="lnkbtnExcel_2007_ColorfulMatrix_Click">带有颜色矩阵的Excel</asp:LinkButton>
                                </li>
                                <li>
                                    <asp:LinkButton ID="lnkbtnExcel_2007_ConditionalFormats" runat="server" ForeColor="Blue" OnClick="lnkbtnExcel_2007_ConditionalFormats_Click">带有格式规则的Excel</asp:LinkButton>
                                </li>
                                <li>
                                    <asp:LinkButton ID="lnkbtnExcel_2007_CreateComment" runat="server" ForeColor="Blue" OnClick="lnkbtnExcel_2007_CreateComment_Click">带有批注的Excel</asp:LinkButton>
                                </li>
                                <li>
                                    <asp:LinkButton ID="lnkbtExcel_2007_CreateCustomProperties" runat="server" ForeColor="Blue" OnClick="lnkbtExcel_2007_CreateCustomProperties_Click">创建用户自定义属性的Excel</asp:LinkButton>
                                </li>
                                <li>
                                    <asp:LinkButton ID="lnkbtnExcel_2007_CreateHeaderFooter" runat="server" ForeColor="Blue" OnClick="lnkbtnExcel_2007_CreateHeaderFooter_Click">带有页眉页脚的Excel</asp:LinkButton>
                                </li>
                                <li>
                                    <asp:LinkButton ID="lnkbtnExcel_2007_DataFormats" runat="server" ForeColor="Blue" OnClick="lnkbtnExcel_2007_DataFormats_Click">带有日期格式的Excel</asp:LinkButton>
                                </li>
                                <li>
                                    <asp:LinkButton ID="lnkbtnExcel_2007_FillBackground" runat="server" ForeColor="Blue" OnClick="lnkbtnExcel_2007_FillBackground_Click">带有填充背景的Excel</asp:LinkButton>
                                </li>
                                <li>
                                    <asp:LinkButton ID="lnkbtnExcel_2007_HideColumnAndRow" runat="server" ForeColor="Blue" OnClick="lnkbtnExcel_2007_HideColumnAndRow_Click">带有隐藏行和列的Excel</asp:LinkButton>
                                </li>
                                <li>
                                    <asp:LinkButton ID="lnkbtnExcel_2007_InsertPictures" runat="server" ForeColor="Blue" OnClick="lnkbtnExcel_2007_InsertPictures_Click">带有图片的Excel</asp:LinkButton>
                                </li>
                                <li>
                                    <asp:LinkButton ID="lnkbtnExcel_2007_LineChart" runat="server" ForeColor="Blue" OnClick="lnkbtnExcel_2007_LineChart_Click">带有线型图表的Excel</asp:LinkButton>
                                </li>
                                <li>
                                    <asp:LinkButton ID="lnkbtnExcel_2007_MeringCells" runat="server" ForeColor="Blue" OnClick="lnkbtnExcel_2007_MeringCells_Click">合并单元格</asp:LinkButton>
                                </li>
                                <li>
                                    <asp:LinkButton ID="lnkbtnExcel_2007_LinkedDropDownLists" runat="server" ForeColor="Blue" OnClick="lnkbtnExcel_2007_LinkedDropDownLists_Click">带有下拉框的Excel</asp:LinkButton>
                                </li>
                                <li>
                                    <asp:LinkButton ID="lnkbtnExcel_2007_MonthlySalaryReport" runat="server" ForeColor="Blue" OnClick="lnkbtnExcel_2007_MonthlySalaryReport_Click">带有每月工资报表的Excel</asp:LinkButton>
                                </li>
                                <li>
                                    <asp:LinkButton ID="lnkbtnExcel_2007_PageSetup" runat="server" ForeColor="Blue" OnClick="lnkbtnExcel_2007_PageSetup_Click">Excel 页面设置(没看出来)</asp:LinkButton>
                                </li>
                                <li>
                                    <asp:LinkButton ID="lnkbtnExcel_2007_PrintSetup" runat="server" ForeColor="Blue" OnClick="lnkbtnExcel_2007_PrintSetup_Click">Excel 打印设置</asp:LinkButton>
                                </li>
                                <li>
                                    <asp:LinkButton ID="lnkbtnExcel_2007_ProtectSheet" runat="server" ForeColor="Blue" OnClick="lnkbtnExcel_2007_ProtectSheet_Click">带有保护机制的Excel</asp:LinkButton>
                                </li>
                                <li>
                                    <asp:LinkButton ID="lnkbtnExcel_2007_ScatterChart" runat="server" ForeColor="Blue" OnClick="lnkbtnExcel_2007_ScatterChart_Click">带有散点图的Excel</asp:LinkButton>
                                </li>
                                <li>
                                    <asp:LinkButton ID="lnkbtnExcel_2007_SetWidthAndHeight" runat="server" ForeColor="Blue" OnClick="lnkbtnExcel_2007_SetWidthAndHeight_Click">设置Excel的长和宽</asp:LinkButton>
                                </li>
                                <li>
                                    <asp:LinkButton ID="lnkbtnExcel_2007_SplitAndFreezePanes" runat="server" ForeColor="Blue" OnClick="lnkbtnExcel_2007_SplitAndFreezePanes_Click">带有拆分和冻结窗口的Excel</asp:LinkButton>
                                </li>
                                <li>
                                    <asp:LinkButton ID="lnkbtnExcel_2007_WritePerformanceTest" runat="server" ForeColor="Blue" OnClick="lnkbtnExcel_2007_WritePerformanceTest_Click">写性能测试(7W行数据)</asp:LinkButton>
                                </li>
                            </ul>
                        </div>
                    </div>
                </div>
            </div>
            <div style="margin-top: 60px;" id="Word_2007">
                <b style="color: #0000FF">关于Word导出 2007</b>
                <div style="width: 100%; height: auto; border: 1px gray dashed; margin-top: 10px; overflow: hidden">
                    <div style="padding-top: 30px; padding-bottom: 20px;">
                        <span class="auto-style1">*主要是针对Word 2007的基本操作</span>
                        <br />
                        <br />
                        <div class="auto-style2">
                            <ul>
                                <li>
                                    <asp:LinkButton ID="lnkbtnWord_2007_CreateEmptyDocument" runat="server" ForeColor="Blue" OnClick="lnkbtnWord_2007_CreateEmptyDocument_Click">创建一个空的Word文档</asp:LinkButton>
                                </li>
                                <li>
                                    <asp:LinkButton ID="lnkbtnWord_2007_InsertPicturesInWord" runat="server" ForeColor="Blue" OnClick="lnkbtnWord_2007_InsertPicturesInWord_Click">带有图片的Word文档</asp:LinkButton>
                                </li>
                                <li>
                                    <asp:LinkButton ID="lnkbtnWord_2007_SimpleDocument" runat="server" ForeColor="Blue" OnClick="lnkbtnWord_2007_SimpleDocument_Click">创建一个简单的Word文档</asp:LinkButton>
                                </li>
                                <li>
                                    <asp:LinkButton ID="lnkbtnWord_2007_SimpleTable" runat="server" ForeColor="Blue" OnClick="lnkbtnWord_2007_SimpleTable_Click">带有表格的Word文档</asp:LinkButton>
                                </li>
                                <li>
                                    <asp:LinkButton ID="lnkbtnWord_2007_UpdateEmbeddedDoc" runat="server" ForeColor="Blue" OnClick="lnkbtnWord_2007_UpdateEmbeddedDoc_Click">更新Word文档</asp:LinkButton>
                                </li>
                                <li>
                                    <asp:LinkButton ID="lnkbtnWord_2007_CreateBullet" runat="server" ForeColor="Blue" OnClick="lnkbtnWord_2007_CreateBullet_Click">带有序列的Word文档</asp:LinkButton>
                                </li>
                                <li>
                                    <asp:LinkButton ID="lnkbtnWord_2007_CreateTable" runat="server" ForeColor="Blue" OnClick="lnkbtnWord_2007_CreateTable_Click">带有合并单元格的Word文档</asp:LinkButton>
                                </li>
                                 <li>
                                    <asp:LinkButton ID="lnkbtn" runat="server" ForeColor="Blue" OnClick="lnkbtnWord_2007_CreateTable_Click">带有合并单元格的Word文档</asp:LinkButton>
                                </li>
                            </ul>
                        </div>
                    </div>
                </div>
            </div>
            <div style="margin-top: 60px;" id="XML">
                <b style="color: #0000FF">压缩文件 OOXML</b>
                <div style="width: 100%; height: auto; border: 1px gray dashed; margin-top: 10px; overflow: hidden">
                    <div style="padding-top: 30px; padding-bottom: 20px;">
                        <span class="auto-style1">*主要是针对压缩文件的基本操作</span>
                        <br />
                        <br />

                        <div class="auto-style2">
                            <ul>
                                <li>
                                    <asp:LinkButton ID="lnkbtn_CreateFile" runat="server" ForeColor="Blue" OnClick="lnkbtn_CreateFile_Click">创建OOXML文件</asp:LinkButton>
                                </li>
                                <li>
                                    <asp:LinkButton ID="lnkbtn_ModifyFile" runat="server" ForeColor="Blue" OnClick="lnkbtn_ModifyFile_Click">修改OOXML文件</asp:LinkButton>
                                </li>
                            </ul>
                        </div>
                    </div>
                </div>
            </div>
            <div style="margin-top: 60px;" id="POIFS">
                <b style="color: #0000FF">Office 自定义属性</b>
                <div style="width: 100%; height: auto; border: 1px gray dashed; margin-top: 10px; overflow: hidden">
                    <div style="padding-top: 30px; padding-bottom: 20px;">
                        <span class="auto-style1">*主要是针对OLE2/ActiveX文档属性的基本操作</span>
                        <br />
                        <br />
                        <div class="auto-style2">
                            <ul>
                                <li>
                                    <asp:LinkButton ID="lnkbtn_CreateCustomProperties" runat="server" ForeColor="Blue" OnClick="lnkbtn_CreateCustomProperties_Click">创建用户自定义属性</asp:LinkButton>
                                </li>
                                <li>
                                    <asp:LinkButton ID="lnkbtn_CreatePOIFS" runat="server" ForeColor="Blue" OnClick="lnkbtn_CreatePOIFS_Click">创建poifs文件</asp:LinkButton>
                                </li>
                                <li>
                                    <asp:LinkButton ID="lnkbtn_CreatePOIFSFileWithPropertie" runat="server" ForeColor="Blue" OnClick="lnkbtn_CreatePOIFSFileWithPropertie_Click">POIFS文件的创建和性能</asp:LinkButton>
                                </li>
                                <li>
                                    <asp:LinkButton ID="lnkbtn_ReadThumbsDB" runat="server" ForeColor="Blue" OnClick="lnkbtn_ReadThumbsDB_Click">读取.db文件</asp:LinkButton>
                                </li>
                            </ul>
                        </div>
                    </div>
                </div>
            </div>
            <div style="margin-top: 60px;" id="NPOI_1">
                <b style="color: #0000FF">NPOI 1.2.*说明</b>
                <div style="width: 100%; height: auto; border: 1px gray dashed; margin-top: 10px; overflow: hidden">
                    <div style="padding-top: 30px; padding-bottom: 20px;">
                        <span class="auto-style1">*主要是针对NPOI 1.2.*版本的说明</span>
                        <br />
                        <br />
                        <div class="auto-style2">
                            <table class="table" style="width: 100%">
                                <tr>
                                    <td colspan="2">NPOI 1.2.x主要由POIFS、DDF、HPSF、HSSF、SS、Util六部分组成。</td>
                                </tr>
                                <tr>
                                    <td>NPOI.POIFS</td>
                                    <td>OLE2/ActiveX文档属性读写库</td>
                                </tr>
                                <tr>
                                    <td>NPOI.DDF</td>
                                    <td>Microsoft Office Drawing读写库</td>
                                </tr>
                                <tr>
                                    <td>NPOI.HPSF</td>
                                    <td>OLE2/ActiveX文档读写库</td>
                                </tr>
                                <tr>
                                    <td>NPOI.HSSF</td>
                                    <td>Microsoft Excel BIFF(Excel 97-2003)格式读写库</td>
                                </tr>
                                <tr>
                                    <td>NPOI.SS</td>
                                    <td>Excel公用接口及Excel公式计算引擎</td>
                                </tr>
                                <tr>
                                    <td>NPOI.Util</td>
                                    <td>基础类库，提供了很多实用功能，可用于其他读写文件格式项目的开发</td>
                                </tr>
                            </table>
                        </div>
                    </div>
                </div>
            </div>
            <div style="margin-top: 60px;" id="NPOI_2">
                <b style="color: #0000FF">NPOI 2.*说明</b>
                <div style="width: 100%; height: auto; border: 1px gray dashed; margin-top: 10px; overflow: hidden">
                    <div style="padding-top: 30px; padding-bottom: 20px;">
                        <span class="auto-style1">*主要是针对NPOI 2.*版本的说明</span>
                        <br />
                        <br />
                        <div class="auto-style2">
                            <table class="table" style="width: 100%">
                                <tr>
                                    <td colspan="3">NPOI 2.*主要由SS, HPSF, DDF, HSSF, XWPF, XSSF, OpenXml4Net, OpenXmlFormats组成，具体列表如下：</td>
                                </tr>
                                <tr>
                                    <td>Assembly名称</td>
                                    <td>模块/命名空间</td>
                                    <td>说明</td>
                                </tr>
                                <tr>
                                    <td>NPOI.DLL</td>
                                    <td>NPOI.POIFS</td>
                                    <td>OLE2/ActiveX文档属性读写库</td>
                                </tr>
                                <tr>
                                    <td>NPOI.DLL</td>
                                    <td>NPOI.DDF</td>
                                    <td>微软Office Drawing读写库</td>
                                </tr>
                                <tr>
                                    <td>NPOI.DLL</td>
                                    <td>NPOI.HPSF</td>
                                    <td>OLE2/ActiveX文档读写库</td>
                                </tr>
                                <tr>
                                    <td>NPOI.DLL</td>
                                    <td>NPOI.HSSF</td>
                                    <td>微软Excel BIFF(Excel 97-2003, doc)格式读写库</td>
                                </tr>
                                <tr>
                                    <td>NPOI.DLL</td>
                                    <td>NPOI.SS</td>
                                    <td>Excel公用接口及Excel公式计算引擎</td>
                                </tr>
                                <tr>
                                    <td>NPOI.DLL</td>
                                    <td>NPOI.Util</td>
                                    <td>基础类库，提供了很多实用功能，可用于其他读写文件格式项目的开发</td>
                                </tr>
                                <tr>
                                    <td>NPOI.OOXML.DLL</td>
                                    <td>NPOI.XSSF</td>
                                    <td>Excel 2007(xlsx)格式读写库</td>
                                </tr>
                                <tr>
                                    <td>NPOI.OOXML.DLL</td>
                                    <td>NPOI.XWPF</td>
                                    <td>Word 2007(docx)格式读写库</td>
                                </tr>
                                <tr>
                                    <td>NPOI.OpenXml4Net.DLL</td>
                                    <td>NPOI.OpenXml4Net</td>
                                    <td>OpenXml底层zip包读写库</td>
                                </tr>
                                <tr>
                                    <td>NPOI.OpenXmlFormats.DLL</td>
                                    <td>NPOI.OpenXmlFormats</td>
                                    <td>微软Office OpenXml对象关系库</td>
                                </tr>
                            </table>
                        </div>
                    </div>
                </div>
            </div>
            <div style="height: 250px">
                <%--只是为了撑开页面高度--%>
            </div>
        </div>
    </form>
</body>
</html>

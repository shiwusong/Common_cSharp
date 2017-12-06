using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;
using System.Windows;
using System.IO;
using Microsoft.Win32;
using System.Diagnostics;

namespace Common
{
    class ExportWord
    {
        //生成word程序对象
        Word.Application wordApp = new Word.Application();
        //生成documnet对象
        Word.Document wordDoc = new Word.Document();
        private Word.Range allRange;//////////////

        public int OpenWord(string templateFile, string fileName)
        {
            //模板文件
            string TemplateFile = templateFile;
            //生成的具有模板样式的新文件
            //模板文件拷贝到新文件
            try
            {
                if (File.Exists(fileName))
                {
                    File.Delete(fileName);
                    File.Copy(TemplateFile, fileName);
                }
                else
                {
                    File.Copy(TemplateFile, fileName);
                }
            }
            catch
            {
                return 0;
            }


            object Obj_FileName = fileName;
            object Visible = false;
            object ReadOnly = false;
            object missing = System.Reflection.Missing.Value;

            //打开文件
            wordDoc = wordApp.Documents.Open(ref Obj_FileName, ref missing, ref ReadOnly, ref missing,
                ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref Visible,
                ref missing, ref missing, ref missing,
                ref missing);
            wordDoc.Activate();
            allRange = wordDoc.Range(missing, missing);///////////////////////
            return 1;

        }
        public void CloseWord()
        {
            object missing = System.Reflection.Missing.Value;
            object IsSave = true;
            wordDoc.Close(ref IsSave, ref missing, ref missing);
            //wordApp.ActiveDocument.Close(ref missing, ref missing, ref missing);
            wordApp.Quit(ref missing, ref missing, ref missing);

        }

        #region 清除word进程
        /**/
        /// <summary>
        /// 清楚word进程
        /// </summary>
        public void KillWordProcess()
        {
          

        }
        #endregion
        //在书签处插入值
        public bool InsertValue(string bookmark, string value)
        {
            object bkObj = bookmark;
            if (wordApp.ActiveDocument.Bookmarks.Exists(bookmark))
            {
                wordApp.ActiveDocument.Bookmarks.get_Item(ref bkObj).Select();
                wordApp.Selection.TypeText(value);
                return true;
            }
            return false;

        }
        //插入表格,bookmark书签
        public Microsoft.Office.Interop.Word.Table InsertTable(string bookmark, int rows, int columns, float width)
        {
            object miss = System.Reflection.Missing.Value;
            object oStart = bookmark;
            Microsoft.Office.Interop.Word.Range range = wordDoc.Bookmarks.get_Item(ref oStart).Range;//表格插入位置
            Microsoft.Office.Interop.Word.Table newTable = wordDoc.Tables.Add(range, rows, columns, ref miss, ref miss);
            //设置表的格式
            newTable.Borders.Enable = 1;  //允许有边框，默认没有边框(为0时报错，1为实线边框，2、3为虚线边框，以后的数字没试过)
            newTable.Borders.OutsideLineWidth = Microsoft.Office.Interop.Word.WdLineWidth.wdLineWidth050pt;//边框宽度
            if (width != 0)
            {
                newTable.PreferredWidth = width;//表格宽度
            }
            newTable.AllowPageBreaks = false;
            return newTable;
        }
        //合并单元格 表名,开始行号,开始列号,结束行号,结束列号
        public void MergeCell(int n, int row1, int column1, int row2, int column2)
        {
            wordDoc.Content.Tables[n].Cell(row1, column1).Merge(wordDoc.Content.Tables[n].Cell(row2, column2));
        }
        //设置表格内容对齐方式 Align水平方向，Vertical垂直方向(左对齐，居中对齐，右对齐分别对应Align和Vertical的值为-1,0,1)
        public void SetParagraph_Table(Microsoft.Office.Interop.Word.Table table, int Align, int Vertical)
        {
            switch (Align)
            {
                case -1: table.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft; break;//左对齐
                case 0: table.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter; break;//水平居中
                case 1: table.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight; break;//右对齐
            }
            switch (Vertical)
            {
                case -1: table.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalTop; break;//顶端对齐
                case 0: table.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter; break;//垂直居中
                case 1: table.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom; break;//底端对齐
            }
        }
        //设置表格字体
        public void SetFont_Table(Microsoft.Office.Interop.Word.Table table, string fontName, double size)
        {
            if (size != 0)
            {
                table.Range.Font.Size = Convert.ToSingle(size);
            }
            if (fontName != "")
            {
                table.Range.Font.Name = fontName;
            }
        }
        public void SetColWidth_Table(int n, int col, float size)
        {
            wordDoc.Content.Tables[n].Columns[col].Width = size;
        }
        public void SetRowHeight_Table(int n, int row, float size)
        {
            wordDoc.Content.Tables[n].Rows[row].Height = size;
        }


        //是否使用边框,n表格的序号,use是或否
        public void UseBorder(int n, bool use)
        {
            if (use)
            {
                wordDoc.Content.Tables[n].Borders.Enable = 1;  //允许有边框，默认没有边框(为0时报错，1为实线边框，2、3为虚线边框，以后的数字没试过)
            }
            else
            {
                wordDoc.Content.Tables[n].Borders.Enable = 2;  //允许有边框，默认没有边框(为0时报错，1为实线边框，2、3为虚线边框，以后的数字没试过)
            }
        }


        //给表格插入一行,n表格的序号从1开始记
        public void AddRow(int n)
        {
            object miss = System.Reflection.Missing.Value;
            wordDoc.Content.Tables[n].Rows.Add(ref miss);
        }


        //给表格添加一行
        public void AddRow(Microsoft.Office.Interop.Word.Table table)
        {
            object miss = System.Reflection.Missing.Value;
            table.Rows.Add(ref miss);
        }

        //给表格插入rows行,n为表格的序号
        public void AddRow(int n, int rows)
        {
            object miss = System.Reflection.Missing.Value;
            Microsoft.Office.Interop.Word.Table table = wordDoc.Content.Tables[n];
            for (int i = 0; i < rows; i++)
            {
                table.Rows.Add(ref miss);
            }
        }


        //给表格中单元格插入元素，table所在表格，row行号，column列号，value插入的元素
        public void InsertCell(Microsoft.Office.Interop.Word.Table table, int row, int column, string value)
        {
            table.Cell(row, column).Range.Text = value;
        }

        //给表格中单元格插入元素，n表格的序号从1开始记，row行号，column列号，value插入的元素
        public void InsertCell(int n, int row, int column, string value)
        {
            wordDoc.Content.Tables[n].Cell(row, column).Range.Text = value;
        }

        //给表格插入一行数据，n为表格的序号，row行号，columns列数，values插入的值
        public void InsertCell(int n, int row, int columns, string[] values)
        {
            Microsoft.Office.Interop.Word.Table table = wordDoc.Content.Tables[n];
            for (int i = 0; i < columns; i++)
            {
                table.Cell(row, i + 1).Range.Text = values[i];
            }
        }


        //插入图片
        public void InsertPicture(string bookmark, string picturePath, float width, float hight, int PictureStart)
        {
            object miss = System.Reflection.Missing.Value;
            object oStart = bookmark;
            Object linkToFile = false;       //图片是否为外部链接
            Object saveWithDocument = true;  //图片是否随文档一起保存 
            object range = wordDoc.Bookmarks.get_Item(ref oStart).Range;//图片插入位置

            wordDoc.InlineShapes.AddPicture(picturePath, ref linkToFile, ref saveWithDocument, ref range);
            wordDoc.Application.ActiveDocument.InlineShapes[PictureStart].Width = width;   //设置图片宽度
            wordDoc.Application.ActiveDocument.InlineShapes[PictureStart].Height = hight;  //设置图片高度
        }



        //插入一段文字,text为文字内容
        public void InsertText(string bookmark, string text, int intSpaceBefore, int intSpaceAfter, int fontSize, int fontBold, string familyName, Word.WdParagraphAlignment align)
        {
            object oStart = bookmark;
            object range = wordDoc.Bookmarks.get_Item(ref oStart).Range;
            Microsoft.Office.Interop.Word.Paragraph wp = wordDoc.Content.Paragraphs.Add(ref range);
            wp.Format.SpaceBefore = intSpaceBefore;
            wp.Range.Text = text;
            wp.Format.SpaceAfter = intSpaceAfter;
            wp.Range.Font.Size = fontSize;
            wp.Range.Font.Bold = fontBold;
            wp.Range.Font.Name = familyName;
            wp.Range.ParagraphFormat.Alignment = align;
            //wp.Range.InsertParagraphAfter();                                      //..........................................................插入段落标记
            //wordDoc.Paragraphs.Last.Range.Text = "\n";
        }
        public void AddContent(string bookmark)
        {
            Object oTrue = true;
            Object oFalse = false;
            object miss = System.Reflection.Missing.Value;
            object oStart = bookmark;
            Word.Range myRange = wordDoc.Bookmarks.get_Item(ref oStart).Range;//目录插入位置
            Object oUpperHeadingLevel = "1";
            Object oLowerHeadingLevel = "2";
            Object oTOCTableID = "目录";

            wordDoc.TablesOfContents.Add(myRange, ref oTrue, ref oUpperHeadingLevel,
                ref oLowerHeadingLevel, ref miss, ref miss, ref oTrue,
                ref oTrue, ref miss, ref oTrue, ref oTrue, ref oTrue);
            wordApp.ActiveDocument.TablesOfContents[1].TabLeader = Word.WdTabLeader.wdTabLeaderMiddleDot;
            wordApp.ActiveDocument.TablesOfContents.Format = Word.WdTocFormat.wdTOCFormal;
            wordDoc.TablesOfContents[1].UpdatePageNumbers(); //更新页码 



        }
        public void FormatContent(int FirstParaArray, int FontSize)
        {
            Microsoft.Office.Interop.Word.TableOfContents myContent = wordDoc.TablesOfContents[1]; //目录  
            Microsoft.Office.Interop.Word.Paragraphs myParagraphs = myContent.Range.Paragraphs; //目录里的所有段，一行一段  
            for (int i = 1; i <= FirstParaArray; i++)
            {
                myParagraphs[i].Range.ParagraphFormat.SpaceBefore = 0; //段前  
                myParagraphs[i].Range.ParagraphFormat.SpaceAfter = 0; //段后间距  
                //myParagraphs[i].Range.Font.Name = "宋体"; //字体  
                myParagraphs[i].Range.Font.Size = FontSize; //小四 
            }
        }
        public void AddPageHeaderFooter(string HeaderFooterText)
        {
            object Nothing = System.Reflection.Missing.Value;

            Microsoft.Office.Interop.Word.WdStatistic stat = Word.WdStatistic.wdStatisticPages;
            int num = wordDoc.ComputeStatistics(stat, ref  Nothing);

            ////添加页眉方法一：
            //WordApp.ActiveWindow.View.Type = WdViewType.wdOutlineView;
            //WordApp.ActiveWindow.View.SeekView = WdSeekView.wdSeekPrimaryHeader;
            //WordApp.ActiveWindow.ActivePane.Selection.InsertAfter( "**公司" );//页眉内容

            ////添加页眉方法二：

            if (wordApp.ActiveWindow.ActivePane.View.Type == Word.WdViewType.wdNormalView ||
            wordApp.ActiveWindow.ActivePane.View.Type == Word.WdViewType.wdOutlineView)
            {
                wordApp.ActiveWindow.ActivePane.View.Type = Word.WdViewType.wdPrintView;
            }
            wordApp.ActiveWindow.View.SeekView = Word.WdSeekView.wdSeekCurrentPageHeader;
            wordApp.Selection.HeaderFooter.LinkToPrevious = false;
            wordApp.Selection.HeaderFooter.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            wordApp.Selection.HeaderFooter.Range.Text = HeaderFooterText;

            wordApp.ActiveWindow.View.SeekView = Word.WdSeekView.wdSeekCurrentPageFooter;
            wordApp.Selection.HeaderFooter.LinkToPrevious = false;
            wordApp.Selection.HeaderFooter.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            wordApp.ActiveWindow.ActivePane.Selection.InsertAfter("");

            //app.Selection.Sections[1].Headers[0].PageNumbers.StartingNumber = 3;
            //app.Selection.Sections[1].Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].PageNumbers.Add(ref   oAlignment, ref   oFirstPage);
            //app.Selection.Sections[1].Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].PageNumbers.NumberStyle = WdPageNumberStyle.wdPageNumberStyleNumberInDash;
            //跳出页眉页脚设置
            wordApp.ActiveWindow.View.SeekView = Word.WdSeekView.wdSeekMainDocument;

        }

        public int PageCount//////////////
        {
            get
            {
                object oMissing = System.Reflection.Missing.Value;
                int pageCount = wordDoc.ComputeStatistics(Word.WdStatistic.wdStatisticPages, oMissing);
                return pageCount;
            }
        }

        public Word.Range GetPages(int pageIndex)/////////////////////////
        {
            object objWhat = Word.WdGoToItem.wdGoToPage;
            object objWhich = Word.WdGoToDirection.wdGoToAbsolute;
            object oMissing = System.Reflection.Missing.Value;
            object objPage = pageIndex;
            Word.Range range1 = wordDoc.GoTo(ref objWhat, ref objWhich, ref objPage, ref oMissing);
            Word.Range range2 = range1.GoToNext(Word.WdGoToItem.wdGoToPage);
            object objStart = range1.Start;
            object objEnd = range2.Start;
            if (range1.Start == range2.Start)
                objEnd = wordDoc.Characters.Count;
            return wordDoc.Range(ref objStart, ref objEnd);
        }


        public void AddPageHeaderFooterFen(string StrHeader)
        {
            object Nothing = System.Reflection.Missing.Value;
            //设置分节符
            for (int i = 1; i <= 2; i++)
            {
                Word.Range range = GetPages(i);
                object oCollapseEnd = Word.WdCollapseDirection.wdCollapseEnd;
                object oPageBreak = Word.WdBreakType.wdSectionBreakContinuous;//分页符   
                range.Collapse(ref oCollapseEnd);
                range.InsertBreak(ref oPageBreak);
                range.Collapse(ref oCollapseEnd);
            }
            wordApp.ActiveWindow.View.Type = Word.WdViewType.wdOutlineView;

            //设置各节的页眉
            wordApp.ActiveWindow.View.SeekView = Word.WdSeekView.wdSeekPrimaryHeader;
            //取消连接到上一页
            foreach (Word.Section section in allRange.Sections)
            {
                section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].LinkToPrevious = false;
            }
            foreach (Word.Section section in allRange.Sections)
            {
                section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text = StrHeader;
                ////设置边框样式
                //section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Paragraphs.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleThinThickThinLargeGap;
                //section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Paragraphs.Borders[Word.WdBorderType.wdBorderTop].LineStyle = Word.WdLineStyle.wdLineStyleNone;
                //section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Paragraphs.Borders[Word.WdBorderType.wdBorderLeft].LineStyle = Word.WdLineStyle.wdLineStyleNone;
                //section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Paragraphs.Borders[Word.WdBorderType.wdBorderRight].LineStyle = Word.WdLineStyle.wdLineStyleNone;
            }
            allRange.Sections[1].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text = "";

            ///////////////////////////////////生成页脚
            wordApp.ActiveWindow.View.SeekView = Word.WdSeekView.wdSeekPrimaryFooter;
            foreach (Word.Section section in allRange.Sections)
            {
                section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].LinkToPrevious = false;
            }
            foreach (Word.Section section in allRange.Sections)
            {

                //section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text = "第";
                //section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].PageNumbers.Add(Word.WdPageNumberAlignment.wdAlignPageNumberCenter, true);
                //section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text ="页   共" + PageCount + "页";
            }
            allRange.Sections[1].Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text = "";
            allRange.Sections[2].Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text = "";
            //allRange.Sections[3].Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text = "";
            //allRange.Sections[4].Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text = "";
            wordApp.ActiveWindow.View.SeekView = Word.WdSeekView.wdSeekMainDocument;
        }
        public void SetPageHeader(string context)
        {
            wordApp.ActiveWindow.View.Type = Word.WdViewType.wdOutlineView;
            wordApp.ActiveWindow.View.SeekView = Word.WdSeekView.wdSeekPrimaryHeader;
            wordApp.ActiveWindow.ActivePane.Selection.InsertAfter(context);
            wordApp.Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            //跳出页眉设置   
            wordApp.ActiveWindow.View.SeekView = Word.WdSeekView.wdSeekMainDocument;
        }

        public void SetBreak()  //插入分节符
        {
            object Nothing = System.Reflection.Missing.Value;
            wordApp.Selection.InsertBreak(ref Nothing);
        }
        /// <summary>
        /// 插入页脚
        /// </summary>
        /// <param name="text">页脚文本</param>
        /// <param name="font">页脚字体</param>
        /// <param name="alignment">显示方式</param>
        public void InsertPageFooter(string text)
        {
            try
            {
                wordApp.ActiveWindow.View.SeekView = Word.WdSeekView.wdSeekCurrentPageFooter;
                wordApp.Selection.HeaderFooter.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                wordApp.Selection.InsertAfter(text);
            }
            catch
            {
                MessageBox.Show("写入失败！");
            }
        }

        /// <summary>
        /// 插入页码
        /// </summary>
        /// <param name="font">页码字体</param>
        /// <param name="alignment">显示方式</param>
        public void InsertPageFooterNumber()
        {
            object Nothing = System.Reflection.Missing.Value;
            try
            {
                wordApp.ActiveWindow.View.SeekView = Word.WdSeekView.wdSeekPrimaryHeader;
                wordApp.Selection.WholeStory();
                wordApp.Selection.ParagraphFormat.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleNone;
                wordApp.ActiveWindow.View.SeekView = Word.WdSeekView.wdSeekMainDocument;
                wordApp.ActiveWindow.View.SeekView = Word.WdSeekView.wdSeekPrimaryFooter;
                wordApp.Selection.TypeText("第");
                object page = Word.WdFieldType.wdFieldPage;
                wordApp.Selection.Fields.Add(wordApp.Selection.Range, ref page, ref Nothing, ref Nothing);
                wordApp.Selection.TypeText("页  共");
                object pages = Word.WdFieldType.wdFieldNumPages;
                wordApp.Selection.Fields.Add(wordApp.Selection.Range, ref pages, ref Nothing, ref Nothing);
                wordApp.Selection.TypeText("页");
                wordApp.ActiveWindow.View.SeekView = Word.WdSeekView.wdSeekMainDocument;
            }
            catch
            {
                MessageBox.Show("写入失败！");
            }
        }
        public void InsertBookMark(string bookmark)
        {
            Microsoft.Office.Interop.Word.Range r = wordApp.Selection.Range;
            object rng = (object)r;
            wordApp.ActiveDocument.Bookmarks.Add(bookmark, ref rng);
            wordApp.ActiveDocument.Bookmarks.DefaultSorting = 0;
            wordApp.ActiveDocument.Bookmarks.ShowHidden = false;
        }
        public void MoveParagraph(int lineNum)
        {
            Object oMissing = System.Reflection.Missing.Value;
            //object WdStory = Word.WdUnits.wdStory;
            //wordApp.Selection.EndKey(ref WdStory, ref oMissing);//定位到文档的最后
            object count = lineNum;
            object WdLine = Microsoft.Office.Interop.Word.WdUnits.wdParagraph;
            //wordApp.Selection.Move(ref?WdLine, ref?count);//移动焦点
            //wordApp.Selection.MoveUp(ref?WdLine, ref?count, oMissing);//移动焦点
            //wordApp.Selection.MoveEnd(ref?WdLine, ref?count);//移动焦点
            //wordApp.Selection.TypeParagraph();//插入段落
        }
        public void MoveRow(int lineNum)
        {
            object count = lineNum;
            object WdLine = Microsoft.Office.Interop.Word.WdUnits.wdLine;//换一行
            wordApp.Selection.Move(ref WdLine, ref count);
        }
        public void MoveCharacter(int lineNum)
        {
            object count = lineNum;
            object wdCharacter = Microsoft.Office.Interop.Word.WdUnits.wdCharacter;
            wordApp.Selection.MoveRight(ref wdCharacter, ref count);
        }
        public void ToNextLine()
        {
            wordApp.Selection.TypeParagraph();
        }
        /// <summary>
        /// 当前位置处插入文字
        /// </summary>
        /// <param name="context">文字内容</param>
        /// <param name="fontSize">字体大小</param>
        /// <param name="fontColor">字体颜色</param>
        /// <param name="fontBold">粗体</param>
        /// <param name="familyName">字体</param>
        /// <param name="align">对齐方向</param>
        public void InsertTextStyle(string context, int fontSize, Word.WdColor fontColor, int fontBold, string familyName, Word.WdParagraphAlignment align)
        {
            //设置字体样式以及方向   
            wordApp.Application.Selection.Font.Size = fontSize;
            wordApp.Application.Selection.Font.Bold = fontBold;
            wordApp.Application.Selection.Font.Color = fontColor;
            wordApp.Selection.Font.Name = familyName;
            wordApp.Application.Selection.ParagraphFormat.Alignment = align;
            wordApp.Application.Selection.TypeText(context);
        }

        #region 判断系统是否装word
        /**/
        /// <summary>
        /// 判断系统是否装word
        /// </summary>
        /// <returns></returns>
        public static bool IsInstallWord()
        {
            RegistryKey machineKey = Registry.LocalMachine;
            if (IsInstallWordByVersion("15.0", machineKey))
            {
                return true;
            }
            if (IsInstallWordByVersion("12.0", machineKey))
            {
                return true;
            }
            if (IsInstallWordByVersion("11.0", machineKey))
            {
                return true;
            }
            return false;
        }
        #endregion



        #region 判断系统是否装某版本的word
        /**/
        /// <summary>
        /// </summary>
        /// <param name="strVersion">版本号</param>
        /// <param name="machineKey"></param>
        /// <returns></returns>
        private static bool IsInstallWordByVersion(string strVersion, RegistryKey machineKey)
        {
            try
            {
                RegistryKey installKey = machineKey.OpenSubKey("Software").OpenSubKey("Microsoft").OpenSubKey("Office").OpenSubKey(strVersion).OpenSubKey("Word").OpenSubKey("InstallRoot");
                if (installKey == null)
                {
                    return false;
                }
                return true;
            }
            catch
            {
                return false;
            }
        }
        #endregion

        #region 查询注册表，判断本机是否安装Office2003,2007和WPS
        public int ExistsRegedit()
        {
            int ifused = 0;
            RegistryKey rk = Registry.LocalMachine;

            //查询Office2003
            RegistryKey f03 = rk.OpenSubKey(@"SOFTWARE\Microsoft\Office\11.0\Word\InstallRoot\");

            //查询Office2007
            RegistryKey f07 = rk.OpenSubKey(@"SOFTWARE\Microsoft\Office\12.0\Word\InstallRoot\");

            ////查询wps
            RegistryKey f10 = rk.OpenSubKey(@"SOFTWARE\Microsoft\Office\14.0\Word\InstallRoot\");
            //查询office13
            RegistryKey f13 = rk.OpenSubKey(@"SOFTWARE\Microsoft\Office\15.0\Word\InstallRoot\");

            //检查本机是否安装Office2003
            if (f03 != null)
            {
                ifused = 1;
            }

            //检查本机是否安装Office2007

            if (f07 != null)
            {
                ifused = 2;
            }

            //检查本机是否安装wps
            if (f10 != null)
            {
                ifused = 3;
            }
            //检查本机是否安装Office2013
            if (f13 != null)
            {
                ifused = 4;
            }

            return ifused;
        }
        #endregion
    }
}


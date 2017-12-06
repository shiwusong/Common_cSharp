using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Windows;
using System.ComponentModel;
using System.IO;
using System.Globalization;
using WpfApp2;
using System.Threading;

namespace Common
{
    class ReportSet
    {
        static public event ShowInfoDelegate ShowInfo;
        public event PropertyChangedEventHandler PropertyChanged;
        double Paer;

        private string strFilePath = AppDomain.CurrentDomain.BaseDirectory + "\\Project.ini";

        private void thread_word(object filepath)
        {
            try
            {
                ExportWord WordReport = new ExportWord();
                string templateFile = App.PATH + "\\report\\WordReport2003.doc";
                //try
                //{
                //    double dblVersion = WordReport.ExistsRegedit();
                //    //模板文件
                //    if (dblVersion == 1)
                //    {
                //        templateFile = App.PATH + "\\TemPic\\report\\WordReport2003.doc";
                //    }
                //    else if (dblVersion == 2)
                //    {
                //        templateFile = App.PATH + "\\TemPic\\report\\WordReport2007.docx";
                //    }
                //    else if(dblVersion == 3)
                //    {
                //        templateFile = App.PATH + "\\TemPic\\report\\WordReport2010.docx";
                //    }
                //    else if(dblVersion == 4)
                //    {
                //        templateFile = App.PATH + "\\TemPic\\report\\WordReport2013.docx";
                //    }
                //}
                //catch
                //{
                //    templateFile = App.PATH + "\\TemPic\\report\\WordReport2003.doc";
                //}

                #region 输入参数
                int  result = WordReport.OpenWord(templateFile, filepath.ToString());
                if (result == 0)
                {
                    ShowInfo("已取消生成Word", true);
                    return;
                }
                try
                {
                    WordReport.InsertValue("rIn", App.wellModel.rIn.ToString());
                    WordReport.InsertValue("r1", App.wellModel.r1.ToString());
                    WordReport.InsertValue("r2", App.wellModel.r2.ToString());
                    WordReport.InsertValue("rOut", App.wellModel.rOut.ToString());
                    WordReport.InsertValue("hWD1", App.wellModel.hWD1.ToString());
                    WordReport.InsertValue("hWD2", App.wellModel.hWD2.ToString());
                    WordReport.InsertValue("hIf", App.wellModel.hIf.ToString());
                    WordReport.InsertValue("hD1", App.wellModel.hD1.ToString());


                    WordReport.InsertValue("v1", App.wellModel.v1.ToString());
                    WordReport.InsertValue("v2", App.wellModel.v2.ToString());
                    WordReport.InsertValue("v3", App.wellModel.v3.ToString());
                    WordReport.InsertValue("e1", App.wellModel.e1.ToString());
                    WordReport.InsertValue("e2", App.wellModel.e2.ToString());
                    WordReport.InsertValue("e3", App.wellModel.e3.ToString());
                    WordReport.InsertValue("lamT1", App.wellModel.lamT1.ToString());
                    WordReport.InsertValue("lamT2", App.wellModel.lamT2.ToString());
                    WordReport.InsertValue("lamT3", App.wellModel.lamT3.ToString());
                    WordReport.InsertValue("afaT1", App.wellModel.afaT1.ToString());
                    WordReport.InsertValue("afaT2", App.wellModel.afaT2.ToString());
                    WordReport.InsertValue("afaT3", App.wellModel.afaT3.ToString());
                    WordReport.InsertValue("epsi1", App.wellModel.epsi1.ToString());
                    WordReport.InsertValue("epsi2", App.wellModel.epsi2.ToString());
                    WordReport.InsertValue("epsi3", App.wellModel.epsi3.ToString());

                    WordReport.InsertValue("pWb", App.wellModel.pWb.ToString());
                    WordReport.InsertValue("pWb2", App.wellModel.pWb2.ToString());

                    WordReport.InsertValue("rhoMud1", App.wellModel.rhoMud1.ToString());
                    WordReport.InsertValue("rhoMud2", App.wellModel.rhoMud2.ToString());
                    WordReport.InsertValue("rhoRock1", App.wellModel.rhoRock1.ToString());
                    WordReport.InsertValue("rhoRock2", App.wellModel.rhoRock2.ToString());
                    WordReport.InsertValue("rhoCem1", App.wellModel.rhoCem1.ToString());
                    WordReport.InsertValue("rhoCem2", App.wellModel.rhoCem2.ToString());
                    WordReport.InsertValue("tTop", App.wellModel.tTop.ToString());
                    WordReport.InsertValue("tBot", App.wellModel.tBot.ToString());
                    WordReport.InsertValue("tFTop", App.wellModel.tFTop.ToString());
                    WordReport.InsertValue("tFBot", App.wellModel.tFBot.ToString());

                }
                catch
                {
                   

                }

                int intpic = 1;
                try
                {
                    string[] ImgTitle = new string[] {
                        App.PATH + "\\TemPic\\井深-内压关系曲线.png", App.PATH + "\\TemPic\\井深-一界面径向应力关系曲线.png",App.PATH + "\\TemPic\\井深-一界面周向应力关系曲线.png",App.PATH + "\\TemPic\\井深-一界面剪切安全系数关系曲线.png",
                        App.PATH + "\\TemPic\\半径-位移关系曲线.png",App.PATH + "\\TemPic\\半径-径向应力关系曲线.png",App.PATH + "\\TemPic\\半径-周向压力关系曲线.png",App.PATH + "\\TemPic\\水泥环半径-径向应力关系曲线.png",App.PATH + "\\TemPic\\水泥环半径-周向应力关系曲线.png",App.PATH + "\\TemPic\\水泥环半径-变形关系曲线.png",App.PATH + "\\TemPic\\套管半径-变形关系曲线.png",
                        App.PATH + "\\TemPic\\塑性区半径关系曲线.png",App.PATH + "\\TemPic\\加载-第一界面径向应力关系曲线.png",App.PATH + "\\TemPic\\加载-第二界面径向应力关系曲线.png",App.PATH + "\\TemPic\\卸载-第一界面径向应力关系曲线.png",App.PATH + "\\TemPic\\卸载-第二界面径向应力关系曲线.png"
                    };
                    string[] Bookmark = new string[] {
                        "img0_0","img0_1","img0_2","img0_3",
                        "img1_0","img1_1","img1_2","img1_3","img1_4","img1_5","img1_6",
                        "img2_0","img2_1","img2_2","img2_3","img2_4"
                    };
                    for (int i = 0; i < ImgTitle.Length; i++)
                    {
                        try
                        {
                            WordReport.InsertPicture(Bookmark[i],ImgTitle[i], 300, 300, intpic);//-------------------------------------------------------井身结构图
                            intpic = intpic + 1;
                        }
                        catch
                        {
                            WordReport.InsertValue(Bookmark[i], "");
                        }
                    }
                }
                catch
                {
                   
                }
                WordReport.CloseWord();
                ShowInfo("已生成弹性Word报告", true);
                WordReport.KillWordProcess();

            }
            catch
            {
            }
                #endregion
        }

        public void ReportOutWord()
        {
            System.Windows.Forms.SaveFileDialog saveFile = new System.Windows.Forms.SaveFileDialog();
            saveFile.AddExtension = true;
            saveFile.Filter = "*.docx|*.docx|all files|*.*";
            saveFile.FilterIndex = 0;
            saveFile.FileName = "井筒完整性分析报告";
            if (saveFile.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string fileName = saveFile.FileName;
                Thread thread = new Thread(new ParameterizedThreadStart(thread_word));
                thread.IsBackground = true;
                thread.Start(fileName);
            }
        }
        public void ReportOutExcel()
        {
            System.Windows.Forms.SaveFileDialog a = new System.Windows.Forms.SaveFileDialog();
            a.Filter = "Excel 工作簿(*.xls;*.xlsx)|*.xls;*.xlsx|All Files|*.*";
            a.FileName = "data.xls";
            a.ShowDialog();
            Thread thread = new Thread(new ParameterizedThreadStart(thread_excel));
            thread.IsBackground = true;
            thread.Start(a.FileName);
        }

        private void thread_excel(object filepath)
        {
            
            //导出到execl  
            try
            {
                //没有数据的话就不往下执行  
                if (App.wellModel.pinEla.Length == 0 && App.wellModel.calhD1.Length == 0)
                {
                    ShowInfo("系统没有数据，退出excel生成", true);
                    return;
                }
                    
                //实例化一个Excel.Application对象  
                Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
                ShowInfo("初始化Excel组件", true);
                try
                {

                    excel.Visible = false;//让后台执行设置为不可见，为true的话会看到打开一个Excel，然后数据在往里写
                    Microsoft.Office.Interop.Excel.Workbook excelWB = excel.Workbooks.Add(System.Type.Missing);
                    Microsoft.Office.Interop.Excel.Worksheet excelWS1 = (Microsoft.Office.Interop.Excel.Worksheet)excelWB.Worksheets[1];
                    Microsoft.Office.Interop.Excel.Worksheet excelWS2 = (Microsoft.Office.Interop.Excel.Worksheet)excelWB.Worksheets[1];

                    //生成Excel中列头名称  
                    excelWS1.Cells[1, 1] = "h";//0
                    excelWS1.Cells[1, 2] = "pIn";//1
                    excelWS1.Cells[1, 3] = "pOut";//2
                    excelWS1.Cells[1, 4] = "p1";//3
                    excelWS1.Cells[1, 5] = "p2";//4
                    excelWS1.Cells[1, 6] = "uSI";//5
                    excelWS1.Cells[1, 7] = "uSO";//6
                    excelWS1.Cells[1, 8] = "uCI";//7
                    excelWS1.Cells[1, 9] = "uCO";//8
                    excelWS1.Cells[1, 10] = "uFI";//9
                    excelWS1.Cells[1, 11] = "uFO";//10
                    excelWS1.Cells[1, 12] = "sigmaRSI";//11
                    excelWS1.Cells[1, 13] = "sigmaThetSI";//12
                    excelWS1.Cells[1, 14] = "sigmaRSO";//13
                    excelWS1.Cells[1, 15] = "sigmaThetSO";//14
                    excelWS1.Cells[1, 16] = "sigmaRCI";//15
                    excelWS1.Cells[1, 17] = "sigmaThetCI";//16
                    excelWS1.Cells[1, 18] = "sigmaRCO";//17
                    excelWS1.Cells[1, 19] = "sigmaThetCO";//18
                    excelWS1.Cells[1, 20] = "sigmaRFI";//19
                    excelWS1.Cells[1, 21] = "sigmaThetFI";//20
                    excelWS1.Cells[1, 22] = "sigmaRFO";//21
                    excelWS1.Cells[1, 23] = "sigmaThetFO";//22
                    ShowInfo("正在写入Excel", true);
                    //存储数据    
                    for (int i = 0; i < 1000; i++)
                    {
                        for (int j = 0; j < 23; j++)
                        {
                            excelWS1.Cells[i + 2, j + 1] = App.wellModel.pinEla[i, j];//第i行第1列                
                        }

                    }            // 按格式保存工作簿



                    excel.ActiveWorkbook.RefreshAll();
                    excel.Workbooks.Application.ActiveWorkbook.RefreshAll();

                    

                    //excelWB.SaveAs(a.FileName, Microsoft.Office.Interop.Excel.XlSaveAction.xlSaveChanges, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange);
                    excelWB.SaveAs(filepath.ToString(), Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal);
                    excelWB.Close(false, null, null);
                }
                catch (Exception ex)
                {
                    ShowInfo("生成失败", true);
                    MessageBox.Show("生成失败：" + ex.Message);

                }

                excel.Quit();
                //excel = null;
                ShowInfo("生成成功", true);
                MessageBox.Show("生成成功！");
            }
            catch (Exception ex)
            {
                ShowInfo("初始化Excel组件失败", true);
                MessageBox.Show(ex.Message, "错误提示");
            }
        }
    }
}

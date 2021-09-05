using System;
using System.Collections.Generic;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using System.Windows.Forms;

namespace ParseTxt2Excel
{
    public partial class Form1 : Form
    {
        string filePath = "";

        /// <summary>
        /// 从模版中读取字段保存下来
        /// </summary>
        List<string> headList = new List<string>();
        /// <summary>
        /// 每个基站对应一个字典结构
        /// </summary>
        List<Dictionary<string, string>> myList = new List<Dictionary<string, string>>();
        public Form1()
        {
            InitializeComponent();
        }
        /// <summary>
        /// 选择文件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private bool openFile()
        {
            ReadExcelTemplate();

            string fileExt = "txt";
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Multiselect = false;//不允许打开多个文件
            dialog.DefaultExt = fileExt;//打开文件时显示的可选文件类型
            dialog.Filter = "txt文件|" + "*." + fileExt ;//打开多个文件
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                filePath = dialog.FileNames[0];
                return true;
            }
            else
            {
                MessageBox.Show("返回文件路径失败");
                return false;
            }
        }
        /// <summary>
        /// 解析文件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e1"></param>
        private void button_start_Click(object sender, EventArgs e1)
        {
            if (!openFile())
                return;
            try
            {
                // 创建一个 StreamReader 的实例来读取文件 
                // using 语句也能关闭 StreamReader
                using (StreamReader sr = new StreamReader(filePath))
                {
                    string line;
                  
                    // 从文件读取并显示行，直到文件的末尾 
                    while ((line = sr.ReadLine()) != null)
                    {
                        if (line.Length == 0 || line.Contains("==") || line.Contains("***") || line.Contains(",") || !line.Contains(":"))
                            continue;
                        
                        if (line.Contains("start:"))
                        {
                            Dictionary<string, string> dictTemp = new Dictionary<string, string>();
                            string[] strs = line.Split(new string[] { "start:" },StringSplitOptions.None);
                            dictTemp["基站IP"] = strs[0].Trim();
                            myList.Add(dictTemp);
                            continue;
                        }

                        string[] strArray = line.Split(':');

                        if (strArray.Length == 2)
                        {
                            getdict()[strArray[0].Trim()] = strArray[1].Trim();
                        }

                        if (strArray.Length > 2)
                        {
                            string v = "";
                            for (int i = 1; i < strArray.Length; i++)
                            {
                                v += strArray[i];
                            }
                            getdict()[strArray[0].Trim()] = v;
                        }

                    }
                }

                WriteToExcel();

                
            }
            catch (Exception e)
            {
                // 显示出错消息
                MessageBox.Show(e.Message);
            }
        }

        private Dictionary<string,string> getdict()
        {
            return myList[myList.Count - 1];
        }
        /// <summary>
        /// 从模版中读取字段保存下来
        /// </summary>
        private void ReadExcelTemplate()
        {
            FileStream fs = null;
        
            IWorkbook workbook = null;
            ISheet sheet = null;

            string filePath = "基站统计模板.xlsx";
            Console.WriteLine(filePath);
            using (fs = File.OpenRead(filePath))
            {
                // 2007版本  
                if (filePath.IndexOf(".xlsx") > 0)
                    workbook = new XSSFWorkbook(fs);
                // 2003版本  
                else if (filePath.IndexOf(".xls") > 0)
                    workbook = new HSSFWorkbook(fs);
                try
                {
                    if (workbook!=null)
                    {
                        sheet = workbook.GetSheetAt(0);//读取第一个sheet，当然也可以循环读取每个sheet 
                        if (sheet!=null)
                        {
                            int rowCount = sheet.LastRowNum;//总行数
                            if (rowCount > 0)
                            {
                                IRow firstRow = sheet.GetRow(sheet.FirstRowNum);//获取第一行
                                int cellCount = firstRow.LastCellNum; //获取列数
                                for (int i = 0; i < cellCount; i++)
                                {
                                    ICell cell = firstRow.GetCell(i);
                                    headList.Add(cell.StringCellValue.Trim()); //字段保存到list
                                }
                              
                            }
                        }
                    }
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message);
                }

            }
    
        }

        private void WriteToExcel()
        {
            FileStream fs = null;

            IWorkbook workbook = null;
            ISheet sheet = null;

            try
            {
                workbook = new HSSFWorkbook();
                sheet = workbook.CreateSheet("Sheet0");//创建一个名称为Sheet0的表  

                // 设置列头
                IRow rowHeader = sheet.CreateRow(0);//excel第一行设为列头  
                for (int c = 0; c < headList.Count; c++)
                {
                    ICell cell = rowHeader.CreateCell(c);
                    cell.SetCellValue(headList[c]);
                }

                for (int i = 0; i < myList.Count; i++)
                {
                    IRow row = sheet.CreateRow(i + 1);
                    for (int j = 0; j < headList.Count; j++)
                    {
                        ICell cell = row.CreateCell(j);
                        if (myList[i].ContainsKey(headList[j]))
                        {
                            cell.SetCellValue(myList[i][headList[j]]);
                        }
                        else
                        {
                            cell.SetCellValue("???");
                        }
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
            DateTime dt = DateTime.Now;
            string name = dt.Year + "年" + dt.Month + "月" + dt.Day + "日" + dt.Hour + "时" + dt.Minute + "分.xls";
            
            using (fs = File.OpenWrite(name))
            {
                workbook.Write(fs);//向打开的这个xls文件中写入数据  
            }
            System.Diagnostics.Process.Start("Explorer.exe", @"/select,"+name);
            //System.Diagnostics.Process open = new System.Diagnostics.Process();
            //open.StartInfo.UseShellExecute = true;
            //open.StartInfo.FileName = name;
            //open.Start();

        }
    }
}

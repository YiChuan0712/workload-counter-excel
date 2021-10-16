using Microsoft.Win32;
using System;
using System.Data;
using System.Reflection;
using System.Windows;
//using System.Windows.Forms;
//using System.Windows.Controls;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using MySql.Data.MySqlClient;



namespace YichuanNET
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow() //构造函数
        {
            InitializeComponent();
        }

        public static string[] files = new string[1024]; //用来保存打开的文件路径
        public static int file_index = 0; //打开文件的个数

        string[,] merged_chart = new string[2048, 32];
        int merged_index = 2; //merge后的行数 初始为2 是因为第一行是表头

        string[,] temp_merged_chart = new string[2048, 32];
        int temp_merged_index = 2;

        Excel.Application name_Excel; //用来保存namesearch之后的excel
        Excel.Workbook name_workbook; //
        Excel.Worksheet name_worksheet; //
        int name_index = 2; //namesearch后的行数 初始为2 是因为第一行是表头

        Excel.Application final_Excel; //用来保存namesearch之后的excel
        Excel.Workbook final_workbook; //
        Excel.Worksheet final_worksheet; //
        int final_index = 2; //namesearch后的行数 初始为2 是因为第一行是表头

        int colcount = -1;

        int col_name = 1;
        int col_hours = 5;

        //导入文件
        public void Button_Click(object sender, RoutedEventArgs e)
        {
            MergedTable.Items.Clear();

            file_index = 0; //每次点击都会重新初始化
            merged_index = 2; //每次点击都会重新初始化
            System.Windows.Forms.OpenFileDialog ofd = new System.Windows.Forms.OpenFileDialog();
            ofd.Multiselect = true;
            ofd.Title = "请选择文件";
            ofd.Filter = "所有文件|*.*|.xls|*.xls|.xlsx|*.xlsx";
            if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                //
                foreach (string file in ofd.FileNames)
                {
                    files[file_index] = file;
                    file_index++;
                }
            }


            for (int i = 0; i < file_index; i++)
            {

                //创建一个Excel对象实例
                Excel.Application oExcel = new Excel.Application();
                //获取路径
                string filepath = files[i];
                //打开
                Excel.Workbook WB = oExcel.Workbooks.Open(filepath);
                //获取文件名
                string ExcelWorkbookname = WB.Name;
                //获取表格数量 
                int worksheetcount = WB.Worksheets.Count;
                //一个一个表格的遍历
                for (int j = 1; j <= worksheetcount; j++)
                {
                    Excel.Worksheet wks = (Excel.Worksheet)WB.Worksheets[j];
                    //获取表格的名字
                    string firstworksheetname = wks.Name;
                    //获取行列数
                    int rowcount = wks.UsedRange.Rows.Count;
                    colcount = wks.UsedRange.Columns.Count;

                    if (i == 0 && j == 1)
                    {
                        for (int l = 1; l <= colcount; l++)
                        {
                            merged_chart[1, l] = Convert.ToString(((Excel.Range)wks.Cells[1, l]).Value);
                        }
                    }
                    //保存
                    for (int k = 2; k <= rowcount; k++)
                    {
                        for (int l = 1; l <= colcount; l++)
                        {
                            merged_chart[merged_index, l] = Convert.ToString(((Excel.Range)wks.Cells[k, l]).Value);
                        }
                        merged_index++;
                    }
                }

            }

            for (int i = 2; i <= merged_index - 1; i++)
            {
                MergedTable.Items.Add
                    (new
                    {
                        colA = merged_chart[i, 1],
                        colB = merged_chart[i, 2],
                        colC = merged_chart[i, 3],
                        colD = merged_chart[i, 4],
                        colE = merged_chart[i, 5],
                        /*colF = merged_chart[i, 6],
                        colG = merged_chart[i, 7],
                        colH = merged_chart[i, 8],
                        colI = merged_chart[i, 9],
                        colJ = merged_chart[i, 10],
                        colK = merged_chart[i, 11],
                        colL = merged_chart[i, 12],
                        colM = merged_chart[i, 13],
                        colN = merged_chart[i, 14],
                        colO = merged_chart[i, 15],
                        colP = merged_chart[i, 16],
                        colQ = merged_chart[i, 17],
                        colR = merged_chart[i, 18],
                        colS = merged_chart[i, 19],
                        colT = merged_chart[i, 20]*/
                    }
                    );
            }

            for (int i = 2; i <= merged_index - 1; i++)
            {
                string temp = merged_chart[i, 2];
                /*int count = 0;
                for (int j = 0; j < temp.Length; j++)
                {
                    if (temp[j] == ',')
                        count++;
                    if (temp[j] == ';')
                        count++;
                    if (temp[j] == '，')
                        count++;
                    if (temp[j] == '；')
                        count++;
                    if (temp[j] == '、')
                        count++;
                }*/
                //temp_merged_chart[temp_merged_index, 2] 
                //int peoplenum = count + 1;
                string[] list = new string[10];
                temp = temp.Replace(" ", "");
                list = temp.Split(new char[] { ',', ';', '，', '；', '、' }, StringSplitOptions.RemoveEmptyEntries);
                int rank = 1;
                foreach (string tempstr in list)//
                {
                    temp_merged_chart[temp_merged_index, 1] = tempstr;
                    temp_merged_chart[temp_merged_index, 2] = Convert.ToString(rank);
                    temp_merged_chart[temp_merged_index, 3] = merged_chart[i, 1];
                    temp_merged_chart[temp_merged_index, 4] = temp;
                    temp_merged_chart[temp_merged_index, 5] = merged_chart[i, 3];
                    temp_merged_chart[temp_merged_index, 6] = merged_chart[i, 4];
                    temp_merged_chart[temp_merged_index, 7] = merged_chart[i, 5];
                    temp_merged_index++;
                    rank++;
                }

            }

            for (int i = 0; i < 2048; i++)
                for (int j = 0; j < 32; j++)
                {
                    merged_chart[i, j] = temp_merged_chart[i, j];
                    merged_index = temp_merged_index;
                }

        }

        private void TextBox_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            //
        }

        //按姓名导出
        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            NameTable.Items.Clear();

            name_index = 2;

            string name = SearchName.Text;

            name_Excel = new Excel.Application();

            object miss = Missing.Value;
            name_workbook = name_Excel.Workbooks.Add(miss);

            name_worksheet = name_workbook.Worksheets[1];

            //int rowcount = merged_worksheet.UsedRange.Rows.Count;
            //int colcount = merged_worksheet.UsedRange.Columns.Count;
            int rowcount = merged_index - 1;
            //int colcount = 20;

            //保存表头
            for (int i = 1; i <= colcount; i++)
            {
                name_worksheet.Cells[1, i] = merged_chart[1, i];
            }

            int flag = 0; //姓名是否存在flag
            for (int i = 2; i <= rowcount; i++)
            {
                string tempname = merged_chart[i, col_name];
                if (tempname.Equals(name))
                {
                    flag = 1; //想查询的姓名是存在的
                    for (int j = 1; j <= colcount; j++)
                    {
                        name_worksheet.Cells[name_index, j] = merged_chart[i, j];
                    }
                    name_index++;
                }
            }

            if (flag == 1)
            {
                for (int i = 2; i <= name_index - 1; i++)
                {
                    /*double temp = 0;
                    for (int j = 12; j <= 18; j++)
                    {
                        if(name_worksheet.Cells[i, j].Value!=null)
                        {
                            temp += name_worksheet.Cells[i, j].Value;
                        }
                    }*/
                    NameTable.Items.Add
                        (new
                        {
                            colA = name_worksheet.Cells[i, 1].Value,
                            colB = name_worksheet.Cells[i, 2].Value,
                            colC = name_worksheet.Cells[i, 3].Value,
                            colD = name_worksheet.Cells[i, 4].Value,
                            colE = name_worksheet.Cells[i, 5].Value,
                            /*colF = name_worksheet.Cells[i, 6].Value,
                            colG = name_worksheet.Cells[i, 7].Value,
                            colH = name_worksheet.Cells[i, 8].Value,
                            colI = name_worksheet.Cells[i, 9].Value,
                            colJ = name_worksheet.Cells[i, 10].Value,
                            colK = name_worksheet.Cells[i, 11].Value,
                            colL = name_worksheet.Cells[i, 12].Value,
                            colM = name_worksheet.Cells[i, 13].Value,
                            colN = name_worksheet.Cells[i, 14].Value,
                            colO = name_worksheet.Cells[i, 15].Value,
                            colP = name_worksheet.Cells[i, 16].Value,
                            colQ = name_worksheet.Cells[i, 17].Value,
                            colR = name_worksheet.Cells[i, 18].Value,
                            //colS = name_worksheet.Cells[i, 19].Value,
                            colS = temp,
                            colT = name_worksheet.Cells[i, 20].Value*/
                        }
                        );
                }
            }
            name_workbook.Close(false, miss, miss);
            name_workbook = null;





        }

        private void DataGrid_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {

        }

        //统计
        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            CalcuTable.Items.Clear();
            //分别用来保存名字 工作时长 index
            string[] namelist = new string[1024];
            string[] hourlist = new string[1024];
            int listindex = 0;
            //从全部数据中一条一条的看
            for (int i = 2; i <= merged_index - 1; i++)
            {
                //首先从这一行中找到姓名
                string name = merged_chart[i, col_name];
                //看这个姓名是否已经有记录
                int flag = 0;
                for (int j = 0; j < listindex; j++)
                {
                    if (namelist[j].Equals(name))
                    {
                        hourlist[j] += "    "+merged_chart[i, col_hours];
                        flag = 1;
                        break;
                    }
                }
                //这个姓名还没有被记录
                if (flag == 0)
                {
                    namelist[listindex] = name;
                    hourlist[listindex] = merged_chart[i, col_hours];
                    listindex++;
                }
            }

            final_Excel = new Excel.Application();

            object miss = Missing.Value;
            final_workbook = final_Excel.Workbooks.Add(miss);

            final_worksheet = final_workbook.Worksheets[1];

            final_worksheet.Cells[1, 1] = "姓名";
            final_worksheet.Cells[1, 2] = "教学奖励";

            for (int i = 2; i <= listindex + 1; i++)
            {
                final_worksheet.Cells[i, 1] = namelist[i - 2];
                final_worksheet.Cells[i, 2] = hourlist[i - 2];
            }

            for (int i = 2; i <= listindex + 1; i++)
            {
                CalcuTable.Items.Add
                    (new
                    {
                        colA = final_worksheet.Cells[i, 1].Value,
                        colB = final_worksheet.Cells[i, 2].Value,
                    }
                    );
            }

            final_workbook.Close(false, miss, miss);
            final_workbook = null;
        }

        //姓名合并
        private void Button_Click_3(object sender, RoutedEventArgs e)
        {
            string from = FromName.Text;
            string to = ToName.Text;
            for (int i = 2; i <= merged_index - 1; i++)
            {
                if (merged_chart[i, col_name].Equals(from))
                {
                    merged_chart[i, col_name] = to;
                }
            }

        }

        private void ListView_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {

        }

        //搜索姓名后导出
        private void Button_Click_4(object sender, RoutedEventArgs e)
        {
            //NameTable.Items.Clear();

            name_index = 2;

            string name = SearchName.Text;

            name_Excel = new Excel.Application();

            object miss = Missing.Value;
            name_workbook = name_Excel.Workbooks.Add(miss);

            name_worksheet = name_workbook.Worksheets[1];

            //int rowcount = merged_worksheet.UsedRange.Rows.Count;
            //int colcount = merged_worksheet.UsedRange.Columns.Count;
            int rowcount = merged_index - 1;
            //int colcount = 20;

            //保存表头
            //for (int i = 1; i <= colcount; i++)
            {
                //name_worksheet.Cells[1, i] = merged_chart[1, i];
            }
            name_worksheet.Cells[1, 1] = "成果完成人";
            name_worksheet.Cells[1, 2] = "成果完成人排序";
            name_worksheet.Cells[1, 3] ="项目名称";
            name_worksheet.Cells[1, 4] ="成果完成人";
            name_worksheet.Cells[1, 5] ="获奖等级";
            name_worksheet.Cells[1, 6] ="年度";
            name_worksheet.Cells[1, 7] ="备注";


            int flag = 0; //姓名是否存在flag
            for (int i = 2; i <= rowcount; i++)
            {
                string tempname = merged_chart[i, col_name];
                if (tempname.Equals(name))
                {
                    flag = 1; //想查询的姓名是存在的
                    for (int j = 1; j <= colcount; j++)
                    {
                        name_worksheet.Cells[name_index, j] = merged_chart[i, j];
                    }
                    name_index++;
                }
            }

            if (flag == 1)
            {
                System.Windows.Forms.FolderBrowserDialog fbd = new System.Windows.Forms.FolderBrowserDialog();
                fbd.Description = "请选择文件路径";
                string foldpath = "";
                if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    foldpath = fbd.SelectedPath + "\\" + name + ".xlsx";
                }
                name_workbook.SaveAs(foldpath, miss, miss, miss, miss, miss, Excel.XlSaveAsAccessMode.xlNoChange, miss, miss, miss, miss, miss);
            }
            name_workbook.Close(false, miss, miss);
            name_workbook = null;
        }

        //工作量统计的导出
        private void Button_Click_5(object sender, RoutedEventArgs e)
        {
            //分别用来保存名字 工作时长 index
            string[] namelist = new string[1024];
            string[] hourlist = new string[1024];
            int listindex = 0;
            //从全部数据中一条一条的看
            for (int i = 2; i <= merged_index - 1; i++)
            {
                //首先从这一行中找到姓名
                string name = merged_chart[i, col_name];
                //看这个姓名是否已经有记录
                int flag = 0;
                for (int j = 0; j < listindex; j++)
                {
                    if (namelist[j].Equals(name))
                    {
                        hourlist[j] +="   "+ merged_chart[i, col_hours];
                        flag = 1;
                        break;
                    }
                }
                //这个姓名还没有被记录
                if (flag == 0)
                {
                    namelist[listindex] = name;
                    hourlist[listindex] = merged_chart[i, col_hours];
                    listindex++;
                }
            }

            final_Excel = new Excel.Application();

            object miss = Missing.Value;
            final_workbook = final_Excel.Workbooks.Add(miss);

            final_worksheet = final_workbook.Worksheets[1];

            final_worksheet.Cells[1, 1] = "姓名";
            final_worksheet.Cells[1, 2] = "教学奖励";

            for (int i = 2; i <= listindex + 1; i++)
            {
                final_worksheet.Cells[i, 1] = namelist[i - 2];
                final_worksheet.Cells[i, 2] = hourlist[i - 2];
            }

            System.Windows.Forms.FolderBrowserDialog fbd = new System.Windows.Forms.FolderBrowserDialog();
            fbd.Description = "请选择文件路径";
            string foldpath = "";
            if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                foldpath = fbd.SelectedPath + "\\" + "汇总" + ".xlsx";
            }
            final_workbook.SaveAs(foldpath, miss, miss, miss, miss, miss, Excel.XlSaveAsAccessMode.xlNoChange, miss, miss, miss, miss, miss);

            final_workbook.Close(false, miss, miss);
            final_workbook = null;
        }

        //刷新
        private void Button_Click_6(object sender, RoutedEventArgs e)
        {
            MergedTable.Items.Clear();
            for (int i = 2; i <= merged_index - 1; i++)
            {
                MergedTable.Items.Add
                    (new
                    {
                        colA = merged_chart[i, 1],
                        colB = merged_chart[i, 2],
                        colC = merged_chart[i, 3],
                        colD = merged_chart[i, 4],
                        colE = merged_chart[i, 5],
                        colF = merged_chart[i, 6],
                        colG = merged_chart[i, 7],
                        colH = merged_chart[i, 8],
                        colI = merged_chart[i, 9],
                        colJ = merged_chart[i, 10],
                        colK = merged_chart[i, 11],
                        colL = merged_chart[i, 12],
                        colM = merged_chart[i, 13],
                        colN = merged_chart[i, 14],
                        colO = merged_chart[i, 15],
                        colP = merged_chart[i, 16],
                        colQ = merged_chart[i, 17],
                        colR = merged_chart[i, 18],
                        colS = merged_chart[i, 19],
                        colT = merged_chart[i, 20]
                    }
                    );
            }

        }

        //转存到数据库
        private void Button_Click_7(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.OpenFileDialog ofd = new System.Windows.Forms.OpenFileDialog();
            ofd.Multiselect = false;
            ofd.Title = "请选择文件";
            ofd.Filter = "所有文件|*.*|.xls|*.xls|.xlsx|*.xlsx";
            string filepath = "";
            if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                //
                foreach (string file in ofd.FileNames)
                {
                    filepath = file;
                }
            }
            Excel.Application app = new Excel.Application();
            Excel.Workbooks wbs = app.Workbooks;
            Excel.Workbook wb = ((Excel.Workbook)wbs.Open(filepath, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value));
            Excel.Worksheet ws = (Excel.Worksheet)wb.Worksheets[1];

            var newCon = "server=127.0.0.1; port=3306; uid=root; pwd=root; database=yichuan; ";

            MySqlConnection con = new MySqlConnection(newCon);
            if (con != null)
                con.Open();
            else
                ;

            int rownum = ws.UsedRange.Rows.Count;
            int colnum = ws.UsedRange.Columns.Count;

            string filename = System.IO.Path.GetFileNameWithoutExtension(filepath);
            //string query = $"LOAD DATA INFILE \"{new_filepath}\" INTO TABLE yichuan.test;";
            string query = $"CREATE TABLE " +
                filename +
                $"( ";
            for (int i = 1; i < colnum; i++)
            {
                query = query + ws.Cells[1, i].Value + $" VARCHAR(50), ";
            }
            query = query + ws.Cells[1, colnum].Value + $" VARCHAR(50) ";
            query = query + $"); ";



            for (int i = 2; i <= rownum; i++)
            {
                query = query + $" INSERT INTO " + filename + $" ( ";
                for (int j = 1; j < colnum; j++)
                {
                    query = query + ws.Cells[1, j].Value + $" , ";
                }
                query = query + ws.Cells[1, colnum].Value;
                query = query + $" ) " + $" VALUES " + $" ( ";

                for (int j = 1; j < colnum; j++)
                {
                    query = query + "\"" + ws.Cells[i, j].Value + "\",";
                }
                query = query + "\"" + ws.Cells[i, colnum].Value + "\"" + $");";
            }
            MySqlCommand cmd = new MySqlCommand(query, con);
            MySqlDataReader rdr = cmd.ExecuteReader();

            con.Close();

            wb.Close(false, Missing.Value, Missing.Value);
            wb = null;

        }
    }
}


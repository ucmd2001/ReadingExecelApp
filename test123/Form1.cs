using System;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using NPOI;
using NPOI.HSSF;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.OpenXmlFormats.Vml;
using System.IO;
using NPOI.XSSF.UserModel;
using NPOI.Util;
using NPOI.SS.Formula.Functions;


namespace test123
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

    }

        private void Form1_Load(object sender, EventArgs e)
        {
            #region Road Excel
            string filePath = @"C:\Users\abc.xlsx";
            IWorkbook wk = null;
            string extension = Path.GetExtension(filePath);
            FileStream fs = File.OpenRead(filePath);
            if (extension.Equals(".xls"))
            {
                //把xls文件中的數據寫入wk中
                wk = new HSSFWorkbook(fs);
            }
            else
            {
                //把xlsx文件中的數據寫入wk中
                wk = new XSSFWorkbook(fs);
            }
            // 關閉excel
            fs.Close();
            #endregion

            int sheetCount = wk.NumberOfSheets;
            ISheet sheet = wk.GetSheetAt(0);
            int rowCount = sheet.LastRowNum;
            IRow row = sheet.GetRow(1);
            IRow index = sheet.GetRow(0);
            Size = new Size(1024, 700);

            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();
            #region 讀取第一個sheet
            for (int i = 1; i <= 102; i++)
            {
                row = sheet.GetRow(i);  //讀取當前數據
                index = sheet.GetRow(i);
                int x_pos = 150;
                int y_pos = 0;

                if (row.GetCell(1).ToString() != "")
                {
                    System.Windows.Forms.Label lb = new System.Windows.Forms.Label();
                    lb.Text = row.GetCell(1).ToString();
                    set_style(lb,x_pos,y_pos,index ,1); //Set_labelStyle

                    for (int j = 1; j <= 15; j++)
                    {
                        #region Label Add
                        y_pos = 40;
                        System.Windows.Forms.Label label = new System.Windows.Forms.Label();
                        label.Name = "label" + "Rack" + j;
                        label.Text = " Good ";
                        set_style(label, x_pos, y_pos, index, j); //Set_labelStyle
                        #endregion
                    }
                }

            }

            #region Button_Seting
            for (int j = 1; j <= 15; j++)
            {
                int x_pos = 150;
                int y_pos = 0;
                Button btn = new Button();
                y_pos = 40;
                btn.Name = "label" + "Rack" + j;
                btn.Text = "Rack " + j.ToString();
                btn.Location = new System.Drawing.Point(x_pos - 150, ((y_pos) * j));
                btn.Click += button2_Click;
                btn.Width = 150;
                btn.Height = 40;
                // 將按鈕加入Panel
                this.Controls.Add(btn);
            }
            #endregion
            #endregion
            stopWatch.Stop();
            Console.WriteLine(stopWatch.ElapsedMilliseconds);

        }

        private void button2_Click(object sender, EventArgs e)
        {
            Form2 f = new Form2();
            f.Visible = true;
            f.Location = new System.Drawing.Point(100 , 100);
        }

        private void set_style(System.Windows.Forms.Label lb,int x_pos ,int y_pos ,IRow index, int j)
        {
            //lb.Location = new System.Drawing.Point((x_pos * int.Parse(index.GetCell(0).ToString())), y_pos);
            //lb.Location = new System.Drawing.Point((x_pos * int.Parse(index.GetCell(0).ToString())), y_pos);
            //lb.Location = new Point(x_pos * int.Parse(index.GetCell(0).ToString()) , ((y_pos) * j) );
            lb.Size = new System.Drawing.Size(150, 40);
            lb.AutoSize = false;
            lb.TextAlign = ContentAlignment.MiddleCenter;
            lb.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            lb.Location = new System.Drawing.Point((x_pos * int.Parse(index.GetCell(0).ToString())), y_pos * j);
            this.Controls.Add(lb);
        }
    }
}

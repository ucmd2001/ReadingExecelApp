using NPOI.HSSF.UserModel;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace test123
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {

        }

        private void Form2_Load(object sender, EventArgs e)
        {
            this.AutoScroll = true;
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

            int x_pos = 150;
            int y_pos = 40;

            for (int i = 1; i <= 2; i++)
            {
                for (int j = 103; j <= 251; j++)
                {
                    row = sheet.GetRow(j);
                    index = sheet.GetRow(j);
                    if (row.GetCell(1).ToString() != "")   //解決合併儲存格為""的問題
                    {
                        if (j == 103)
                        {
                            System.Windows.Forms.Label lb = new System.Windows.Forms.Label();
                            if(i == 1)
                            {
                                lb.Text = "RB";
                            }
                            else
                            {
                                lb.Text = "Value";
                            }
                            set_style(lb , x_pos , y_pos , i , index) ;
                        }
                        else
                        {
                            System.Windows.Forms.Label lb = new System.Windows.Forms.Label();
                            if (i == 1)
                            {
                                lb.Text = row.GetCell(1).ToString();
                            }
                            else
                            {
                                lb.Text = "Good";
                            }
                            set_style(lb, x_pos, y_pos, i, index);
                        }

                    }
                }
            }
        }
        private void set_style(System.Windows.Forms.Label lb, int x_pos ,int y_pos ,int i , IRow index )
        {
            lb.Size = new System.Drawing.Size(150, 40);
            lb.AutoSize = false;
            lb.TextAlign = ContentAlignment.MiddleCenter;
            lb.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            lb.Location = new System.Drawing.Point((x_pos * i) - 150, y_pos * (int.Parse(index.GetCell(0).ToString()) - 65));
            this.Controls.Add(lb);
        }
    }
}

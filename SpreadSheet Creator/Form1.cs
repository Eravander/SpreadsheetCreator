using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SpreadSheet_Creator
{
    public partial class Form1 : Form
    {
        
        List<string> list;
        public Form1()
        {
            InitializeComponent();
            list = new List<string>();
            listBox1.Items.AddRange(new string[] { "FA19", "SP20" });
            listBox1.ItemCheck += new ItemCheckEventHandler(ListBox1_ItemCheck);
        }

        private void GenBtn_Click(object sender, EventArgs e)
        {
            StringBuilder sb = new StringBuilder();
            foreach (string str in list)
                sb.AppendLine(str);
            MessageBox.Show(sb.ToString());

            Excel ex = new Excel();
            ex.CreateNewFile();
            ex.SaveAs(@"Test");
            ex.Close();

            //ReadData();
            WriteData();

        }


        private void ListBox1_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            string item = listBox1.SelectedItem.ToString();
            if (e.NewValue == CheckState.Checked)
            {
                if (!list.Contains(item))
                    list.Add(item);
            }
            else
            {
                if (list.Contains(item))
                    list.Remove(item);
            }
        }

        private void ListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        public void ReadData()
        {
            Excel ex = new Excel(@"Test", 1);
            string[,] read = ex.ReadRange(1, 1, 2749, 3);
            ex.Close();
        }
        public void WriteData()
        {
            Excel excel = new Excel(@"Test.xlsx", 1);
            excel.WriteToCell(0, 0, "Test2");
            excel.Save();
            excel.SaveAs(@"Test.xlsx");

            excel.Close();
        }
    }
}

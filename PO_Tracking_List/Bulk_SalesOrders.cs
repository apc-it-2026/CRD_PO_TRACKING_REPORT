using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PO_Tracking_List
{
    public partial class Bulk_SalesOrders : Form
    {

        public delegate void DataChangeHandler(object sender, CheckDataChangeEventArgs args);
        public event DataChangeHandler DataChange;
        DataTable dataDt = null;
        string condition1;

       // bool IsCheckBoxClicked = false;

        public Bulk_SalesOrders(DataTable dt, string s1)
        {
            InitializeComponent();

            SJeMES_Framework.Common.UIHelper.UIUpdate(this.Name, this, Program.client, "", Program.client.Language);
            dataDt = dt;
            condition1 = s1;          
            dataGridView1.DataSource = dataDt;

            //CheckBox.MouseClick += new MouseEventHandler(CheckBox_MouseClick);


        }

       

        public void OnDataChange(object sender, CheckDataChangeEventArgs args)
        {
            DataChange?.Invoke(this, args);
        }

        public class CheckDataChangeEventArgs : EventArgs
        {
            public string value1 { get; set; }

            public CheckDataChangeEventArgs(string s1)
            {

                value1 = s1;
            }
        }

        private void btn_SO_Confirm_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count > 0)
            {
                string item_noList = "";
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    if ((bool)row.Cells[0].EditedFormattedValue)
                    {
                        string chk_item_no = row.Cells[1].Value == null ? "" : row.Cells[1].Value.ToString();
                        if (!"".Equals(chk_item_no))
                        {
                            if ("".Equals(item_noList))
                                item_noList += "'" + chk_item_no + "'";
                            else
                                item_noList += ",'" + chk_item_no + "'";
                        }
                    }
                }
                if (!"".Equals(item_noList))
                {
                    OnDataChange(this, new CheckDataChangeEventArgs(item_noList));
                    this.Close();
                    this.Dispose();
                }
                else
                {
                    string msg = SJeMES_Framework.Common.UIHelper.UImsg("Please Tick atleast one data in the List！", Program.client, "", Program.client.Language);
                    MessageBox.Show(msg);
                    return;
                }

            }
            else
            {
                string msg = SJeMES_Framework.Common.UIHelper.UImsg("The list has no data！", Program.client, "", Program.client.Language);
                MessageBox.Show(msg);
                return;
            }
        }

        
        private void textselect_TextChanged(object sender, EventArgs e)
        {
            //dataGridView1.DataSource = null;

            
            string textInfo = textselect.Text;
            string filter = condition1 + " like '%" + textInfo + "%'";
            DataView dv = dataDt.DefaultView;
            dv.RowFilter = filter;
            dataGridView1.DataSource = dv;
           
        }
     

        private void btn_Select_All_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count > 0)
            {
                if (btn_Select_All.Text == "SelectAll")
                {
                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        row.Cells[0].Value = true;
                    }

                    btn_Select_All.Text = "DeSelectAll";
                }
                else
                {
                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        row.Cells[0].Value = false;
                    }
                    btn_Select_All.Text = "SelectAll";
                }
            }
            else
            {
                SJeMES_Control_Library.MessageHelper.ShowErr(this, "No Data Found");
            }
        }


        List<int> ls = new List<int>();

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            //string so = "";
            //string bulk_so = "";


            DataGridViewCheckBoxCell ch1 = new DataGridViewCheckBoxCell();
            ch1 = (DataGridViewCheckBoxCell)dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[0];

            if (ch1.Value == null)
            {
                ch1.Value = false;
            }
            else
            {
                lbl_bulk_item.Text += ch1.Value;
            }

            //for (int i = 0; i <= dataGridView1.Rows.Count - 1; i++)
            //{
            //    if (Convert.ToBoolean(dataGridView1.Rows[i].Cells["chk"].Value) == true)
            //    {
            //        so = dataGridView1.Rows[i].Cells["SALES_ORDER"].Value.ToString();

            //    }
            //    bulk_so += so;
            //}
            //lbl_bulk_item.Text = bulk_so;
        }








    }
}

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MaterialSkin.Controls;
using System.Reflection;

namespace PO_Tracking_List.DoubleClickForm
{
    public partial class CusAssmeblyQtyForm : MaterialForm
    {

        /// <summary>
        /// 
        /// </summary>
        /// <param name="title">form title</param>
        /// <param name="p">health value pair  ===   (tabpage_name,datatable)</param>
        public CusAssmeblyQtyForm(string title,Dictionary<string,DataTable> p )
        {
            InitializeComponent();
            SJeMES_Framework.Common.UIHelper.UIUpdate(this.Name, this, Program.client, "", Program.client.Language);
            this.Text = title;
            try
            {
                foreach (string page_title in p.Keys)
                {
                    this.tabControl1.TabPages.Add(new TabPage(page_title)
                    {
                        Name = page_title
                    });

                    tabControl1.TabPages[page_title].Controls.Add(new DataGridView()
                    {
                        DataSource = p[page_title],
                        Dock = DockStyle.Fill,
                        ReadOnly = true,
                        AllowUserToAddRows = false

                    });
                }
                //SJeMES_Framework.Common.UIHelper.UIUpdate(this.Name, this, Program.client, "", Program.client.Language);
            }
            catch (Exception exp) {
                SJeMES_Control_Library.MessageHelper.ShowErr(this, exp.Message);
            
            }
        }
    }
}

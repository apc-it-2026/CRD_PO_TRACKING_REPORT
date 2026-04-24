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

namespace PO_Tracking_List.DoubleClickForm
{
    public partial class InStockQtyForm : MaterialForm
    {
        string se_id = null;
        string size_no = null;

        public InStockQtyForm(string vSeId, string vSizeNo)
        {
            InitializeComponent();
            se_id = vSeId;
            size_no = vSizeNo;
            SJeMES_Framework.Common.UIHelper.UIUpdate(this.Name, this, Program.client, "", Program.client.Language);
        }

        private void InStockQtyForm_Load(object sender, EventArgs e)
        {
            SJeMES_Framework.Common.UIHelper.UIUpdate(this.Name, this, Program.client, "", Program.client.Language);
            GetInStockQtyDetail(se_id, size_no);
        }

        private void GetInStockQtyDetail(string se_id, string size_no)
        {
            Dictionary<string, Object> p = new Dictionary<string, object>();
            p.Add("vSeId", se_id);
            p.Add("vSizeNo", size_no);
            p.Add("vProcessNo", "C");
            string ret = SJeMES_Framework.WebAPI.WebAPIHelper.Post(Program.client.APIURL, "KZ_SFCAPI_WorkOrder", "KZ_SFCAPI_WorkOrder.Controllers.SFC_Po_Tracking_ListServer", "GetInStockQtyDetail", Program.client.UserToken, Newtonsoft.Json.JsonConvert.SerializeObject(p));
            if (Convert.ToBoolean(Newtonsoft.Json.JsonConvert.DeserializeObject<Dictionary<string, object>>(ret)["IsSuccess"]))
            {
                string json = Newtonsoft.Json.JsonConvert.DeserializeObject<Dictionary<string, object>>(ret)["RetData"].ToString();
                DataTable dt = SJeMES_Framework.Common.JsonHelper.GetDataTableByJson(json);
                if (dt.Rows.Count > 0)
                {
                    dataGridView1.DataSource = dt;
                }
                else
                {
                    string msg = SJeMES_Framework.Common.UIHelper.UImsg("查无此数据！", Program.client, "", Program.client.Language);
                    SJeMES_Control_Library.MessageHelper.ShowErr(this, msg);
                }
            }
            else
            {
                SJeMES_Control_Library.MessageHelper.ShowErr(this, Newtonsoft.Json.JsonConvert.DeserializeObject<Dictionary<string, object>>(ret)["ErrMsg"].ToString());
            }
        }
    }
}

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
using Newtonsoft.Json;
using PO_Tracking_List.DoubleClickForm;
using TExcel = Microsoft.Office.Interop.Excel;

namespace PO_Tracking_List
{
    public partial class PO_Tracking_List : MaterialForm
    {
        Dictionary<int, string> CusFormIndex = new Dictionary<int, string>();
        DataTable dt;
        public PO_Tracking_List()
        {
            InitializeComponent();
            SJeMES_Framework.Common.UIHelper.UIUpdate(this.Name, this, Program.client, "", Program.client.Language);

            this.WindowState = FormWindowState.Maximized;

            CusFormIndex.Add(dataGridView1.Columns["PackingBalQty"].Index, "A");
            CusFormIndex.Add(dataGridView1.Columns["AssemblyBalQty"].Index, "L");
            CusFormIndex.Add(dataGridView1.Columns["CuttingBalQty"].Index, "C");
            CusFormIndex.Add(dataGridView1.Columns["StitchingBalQty"].Index, "S");

            //Workcenter 
            textBox_DeptNo.Visible = false;
        
            // Hiding tabpage3
            tabControl1.TabPages.Remove(tabPage3);
            loadbl12.Visible = false;
            progressBar2.Visible = false;

            #region Datagrid column make hide
            dataGridView1.Columns["CRD_DUE_DAY_STATUS"].Visible = false;
   
         //   dataGridView1.Columns["Column2"].Visible = false;

            dataGridView2.Columns["CRD_DUE_DAY_STATUS_SO"].Visible = false;
            dataGridView2.Columns["cutQty_2"].Visible = false;
            dataGridView2.Columns["stitchingQty_2"].Visible = false;
            dataGridView2.Columns["outSoleQty_2"].Visible = false;
            dataGridView2.Columns["assmeblyQty_2"].Visible = false;
            dataGridView2.Columns["packingQty_2"].Visible = false;
            #endregion

        }

        private void F_BDM_Task_List_Load(object sender, EventArgs e)
        {
            // Item_no_List.Enabled = false;

            this.dataGridView1.AutoGenerateColumns = false;
            this.dataGridView2.AutoGenerateColumns = false;
            this.dataGridView3.AutoGenerateColumns = false;
            this.dateTimePicker5.Value = DateTime.Now.AddDays(1 - DateTime.Now.Day);
            this.dateTimePicker2.Value = DateTime.Now.AddDays(0);
            this.dateTimePicker7.Value = DateTime.Now.AddDays(1 - DateTime.Now.Day);
            this.dateTimePicker4.Value = DateTime.Now.AddDays(0);
            GetDepts();
            GetCountry();

            //  dataGridView1.CellFormatting += dataGridView1_CellFormatting;
            dataGridView1.RowPrePaint += dataGridView1_RowPrePaint;
            dataGridView2.RowPrePaint += dataGridView2_RowPrePaint;

            psddcheckbx.Visible = false;
            label8.Visible = false;
            psddPickerto.Visible = false;
            psddfrompicker.Visible = false; label11.Visible = false;

        }

        //Aded by Ashok on 2026/02/03 to filter based on country  
        private void GetCountry()    
        {
            DataTable dt = null;
            Dictionary<string, Object> p = new Dictionary<string, object>();
            string ret = SJeMES_Framework.WebAPI.WebAPIHelper.Post(Program.client.APIURL, "KZ_QCO", "KZ_QCO.Controllers.MESUpdateServer", "GetCountry", Program.client.UserToken, Newtonsoft.Json.JsonConvert.SerializeObject(p));
            if (Convert.ToBoolean(Newtonsoft.Json.JsonConvert.DeserializeObject<Dictionary<string, object>>(ret)["IsSuccess"]))
            {
                string json = Newtonsoft.Json.JsonConvert.DeserializeObject<Dictionary<string, object>>(ret)["RetData"].ToString();
                dt = SJeMES_Framework.Common.JsonHelper.GetDataTableByJson(json);
                DataRow dr = dt.NewRow();
                dr["Country"] = ""; 
                dt.Rows.InsertAt(dr, 0);
                comboBox2.DataSource = dt;
                comboBox2.DisplayMember = "Country";
                comboBox2.ValueMember = "Country";
            }
            else
            {
                SJeMES_Control_Library.MessageHelper.ShowErr(this, Newtonsoft.Json.JsonConvert.DeserializeObject<Dictionary<string, object>>(ret)["ErrMsg"].ToString());
            }
            
        }
        private DataTable GetDepts()    //Get department information
        {
            var orderSource = new AutoCompleteStringCollection();   //Store database query results
            DataTable dt = null;
            Dictionary<string, Object> p = new Dictionary<string, object>();
            string ret = SJeMES_Framework.WebAPI.WebAPIHelper.Post(Program.client.APIURL, "KZ_MESAPI", "KZ_MESAPI.Controllers.GeneralServer", "GetAllDepts", Program.client.UserToken, Newtonsoft.Json.JsonConvert.SerializeObject(p));
            if (Convert.ToBoolean(Newtonsoft.Json.JsonConvert.DeserializeObject<Dictionary<string, object>>(ret)["IsSuccess"]))
            {
                string json = Newtonsoft.Json.JsonConvert.DeserializeObject<Dictionary<string, object>>(ret)["RetData"].ToString();
                dt = SJeMES_Framework.Common.JsonHelper.GetDataTableByJson(json);

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    orderSource.Add(dt.Rows[i]["DEPARTMENT_CODE"].ToString() + "|" + dt.Rows[i]["DEPARTMENT_NAME"].ToString());
                }
                //【Production department】Bind data source
                textBox_DeptNo.AutoCompleteCustomSource = orderSource;   //bind data source
                textBox_DeptNo.AutoCompleteMode = AutoCompleteMode.Suggest;    //Show related dropdown
                textBox_DeptNo.AutoCompleteSource = AutoCompleteSource.CustomSource;   //set properties
            }
            else
            {
                SJeMES_Control_Library.MessageHelper.ShowErr(this, Newtonsoft.Json.JsonConvert.DeserializeObject<Dictionary<string, object>>(ret)["ErrMsg"].ToString());
            }
            return dt;
        }

        private async void btnSelect_Click(object sender, EventArgs e)
        {
            #region Added by Shyam 2025.08.07  

            if (string.IsNullOrEmpty(PO_txt.Text) && string.IsNullOrEmpty(SOtxt.Text) && (lpdcheck.Checked == false) && (crdcheck.Checked == false) && (string.IsNullOrEmpty(bulktxt.Text)))
            {
                SJeMES_Control_Library.MessageHelper.ShowErr(this, "Please Select Any One Condition PO or SO or CRD or LPD or Bulk SO List !! ");
                return;
            }
            if (!Validate_CRD_Date())
            {
                SJeMES_Control_Library.MessageHelper.ShowErr(this, " CRD Range Must Not Exceed 3 Months!");
                return;
            }
           
            if (!Validate_LPD_Date())
            {
                SJeMES_Control_Library.MessageHelper.ShowErr(this, " LPD Range Must Not Exceed 3 Months!");
                return;
            }

            #endregion

            if (tabControl1.SelectedIndex == 0) //Query tab: Query by order size
            {
                dataGridView1.DataSource = null;

                Cursor.Current = Cursors.WaitCursor;



                Dictionary<string, Object> p = new Dictionary<string, object>();
                p.Add("vSeId", SOtxt.Text);
                p.Add("vPO", PO_txt.Text);                       
                p.Add("SeIdList", bulktxt.Text.Trim().ToString());

                p.Add("vCheckShip", checkBox6.Checked);

                p.Add("vCheckCRD", crdcheck.Checked);
                p.Add("vBeginDate", dateTimePicker5.Value.ToShortDateString());
                p.Add("vEndDate", dateTimePicker6.Value.ToShortDateString());
                p.Add("vCheckLPD", lpdcheck.Checked);
                p.Add("vLPDBeginDate", dateTimePicker7.Value.ToShortDateString());
                p.Add("vLPDEndDate", dateTimePicker8.Value.ToShortDateString());
                // Added by Khaleel 2025/12/10
                string plantValue = string.Empty;

                bool isSelected = plantcm.SelectedItem != null;
                bool isTyped = !string.IsNullOrWhiteSpace(plantcm.Text);


                if (isSelected)
                {
                    plantValue = plantcm.SelectedItem.ToString();
                }
                else
                {
                    //plantValue = plantcombo.Text.Trim().ToUpper();
                    plantValue = plantcm.Text.Trim().ToUpper();
                }

                p.Add("plant", plantValue); 
              
                //Added by Ashok on 2026/02/03
                p.Add("Country", comboBox2.Text);
                p.Add("FGT", fgtcombo.Text); //Added by khaleel on 2026/02/25

                loadbl12.Visible = true;
                progressBar2.Style = ProgressBarStyle.Marquee;
                progressBar2.Visible = true;


                //string ret = SJeMES_Framework.WebAPI.WebAPIHelper.Post(Program.client.APIURL, "KZ_QCO", "KZ_QCO.Controllers.MESUpdateServer", "GetData_P", Program.client.UserToken, Newtonsoft.Json.JsonConvert.SerializeObject(p));
                string ret = await Task.Run(() =>
    SJeMES_Framework.WebAPI.WebAPIHelper.Post(
        Program.client.APIURL,
        "KZ_QCO",
        "KZ_QCO.Controllers.MESUpdateServer",
        "GetData_P",
        Program.client.UserToken,
        Newtonsoft.Json.JsonConvert.SerializeObject(p)
    ));

                #region old Block without optmize

                //if (Convert.ToBoolean(Newtonsoft.Json.JsonConvert.DeserializeObject<Dictionary<string, object>>(ret)["IsSuccess"]))
                //{

                //    string json = Newtonsoft.Json.JsonConvert.DeserializeObject<Dictionary<string, object>>(ret)["RetData"].ToString();
                //    DataTable dtJson = SJeMES_Framework.Common.JsonHelper.GetDataTableByJson(json);
                //    // This is Date Format for Datagrid Cell 
                //    if (dtJson.Columns.Contains("assembly_finished_date"))
                //    {
                //        foreach (DataRow row in dtJson.Rows)
                //        {
                //            var val = row["assembly_finished_date"];

                //            if (val == DBNull.Value || string.IsNullOrWhiteSpace(val.ToString()))
                //            {
                //                row["assembly_finished_date"] = "";   // keep blank
                //            }
                //            else
                //            {
                //                if (DateTime.TryParse(val.ToString(), out DateTime d))
                //                    row["assembly_finished_date"] = d.ToString("yyyy-MM-dd");
                //                else
                //                    row["assembly_finished_date"] = "";  // invalid format protection
                //            }
                //        }

                //        dtJson.Columns["assembly_finished_date"].DataType = typeof(string);
                //    }
                //    if (dtJson.Columns.Contains("PC_DATE"))
                //    {
                //        foreach (DataRow row in dtJson.Rows)
                //        {
                //            var val = row["PC_DATE"];

                //            if (val == DBNull.Value || string.IsNullOrWhiteSpace(val.ToString()))
                //            {
                //                row["PC_DATE"] = "";   // keep blank
                //            }
                //            else
                //            {
                //                if (DateTime.TryParse(val.ToString(), out DateTime d))
                //                    row["PC_DATE"] = d.ToString("yyyy-MM-dd");
                //                else
                //                    row["PC_DATE"] = "";  // invalid format protection
                //            }
                //        }

                //        dtJson.Columns["PC_DATE"].DataType = typeof(string);
                //    }


                //    if (dtJson.Rows.Count == 0)
                //    {
                //        SJeMES_Control_Library.MessageHelper.ShowErr(this, "No such data");
                //        loadbl12.Visible = false;
                //        progressBar2.Visible = false;
                //        return;
                //    }


                //    #region Add Balance Columns (Cutting, Stitching, Assembly, Packing) Added by Khaleel 2025.11.20
                //    void AddColumnIfMissing(string colName)
                //    {
                //        if (!dtJson.Columns.Contains(colName))
                //            dtJson.Columns.Add(colName, typeof(int));
                //    }

                //    AddColumnIfMissing("CuttingBalQty");
                //    AddColumnIfMissing("StitchingBalQty");
                //    AddColumnIfMissing("AssemblyBalQty");
                //    AddColumnIfMissing("PackingBalQty");
                //    AddColumnIfMissing("FGBalqty");
                //    AddColumnIfMissing("ShipBalqty");
                //    #endregion





                //    #region Calculate Balance Qty for EACH ROW (Correct Rule)

                //    foreach (DataRow row in dtJson.Rows)
                //    {
                //        int order = ToIntSafe(row["se_qty"]);
                //        int cut = ToIntSafe(row["cutQty"]);
                //        int stitch = ToIntSafe(row["stitchingQty"]);
                //        int assm = ToIntSafe(row["assmeblyQty"]);
                //        int pack = ToIntSafe(row["packingQty"]);

                //        int FG = ToIntSafe(row["inStock_qty"]);

                //        int ship = ToIntSafe(row["shipping_qty"]);

                //        // Cutting Balance
                //        row["CuttingBalQty"] = Math.Max(order - cut, 0);

                //        // Stitching Balance
                //        row["StitchingBalQty"] = Math.Max(order - stitch, 0);

                //        // Assembly Balance
                //        row["AssemblyBalQty"] = Math.Max(order - assm, 0);

                //        // Packing Balance
                //        row["PackingBalQty"] = Math.Max(order - pack, 0);

                //        // FG Balance
                //        row["FGBalqty"] = Math.Max(order - FG, 0);

                //        // Ship Balance
                //        row["ShipBalqty"] = Math.Max(order - ship, 0);
                //    }

                //    #endregion
                //    #region This Block For Packing Not Finished Then Show the assembly_finished_date Empty
                //    if (dtJson.Columns.Contains("assembly_finished_date"))
                //    {
                //        foreach (DataRow row in dtJson.Rows)
                //        {
                //            // If not fully Pack → hide date
                //            if (ToIntSafe(row["PackingBalQty"]) != 0)
                //            {
                //                row["assembly_finished_date"] = "";
                //            }

                //        }
                //    }
                //    #endregion

                //    #region For check Finished Data Hide

                //    int ToIntSafe(object value)
                //    {
                //        if (value == null || value == DBNull.Value) return 0;

                //        string s = value.ToString().Trim();
                //        if (string.IsNullOrEmpty(s)) return 0;

                //        decimal d;
                //        if (decimal.TryParse(s, out d))
                //            return (int)Math.Round(d);

                //        return 0;
                //    }

                //    bool IsFinishedLine(DataRow r)
                //    {
                //        bool isSupplier = ToIntSafe(r["IS_SUPPLIER_ORDER"]) == 1;

                //        if (isSupplier)
                //        {
                //            // Supplier → Assembly & Packing only
                //            return
                //                ToIntSafe(r["AssemblyBalQty"]) <= 0 &&
                //                ToIntSafe(r["PackingBalQty"]) <= 0;
                //        }
                //        else
                //        {
                //            // Normal → All stages
                //            return
                //                ToIntSafe(r["CuttingBalQty"]) <= 0 &&
                //                ToIntSafe(r["StitchingBalQty"]) <= 0 &&
                //                ToIntSafe(r["AssemblyBalQty"]) <= 0 &&
                //                ToIntSafe(r["PackingBalQty"]) <= 0;
                //        }
                //    }
                //    if (checkBox3.Checked)
                //    {
                //        var removeRows = dtJson.AsEnumerable()
                //            .Where(r =>
                //                !string.IsNullOrWhiteSpace(r["size_no"].ToString()) &&
                //                r["mold_no"].ToString() != "Total" &&
                //                IsFinishedLine(r))
                //            .ToList();

                //        foreach (var row in removeRows)
                //            dtJson.Rows.Remove(row);

                //        dtJson.AcceptChanges();
                //    }

                //    #endregion


                //    #region GROUP TOTAL: Add Total_By_SO for each SE_ID
                //    var seIds = dtJson.AsEnumerable()
                //                      .Select(r => r.Field<string>("se_id"))
                //                      .Distinct()
                //                      .ToList();

                //    foreach (var seId in seIds)
                //    {
                //        var rows = dtJson.AsEnumerable()
                //                         .Where(r => r.Field<string>("se_id") == seId &&
                //                                     r.Field<string>("size_no") != "Total_By_SO" &&
                //                                     r.Field<string>("size_no") != "Total")
                //                         .ToList();

                //        if (!rows.Any()) continue;

                //        DataRow totalBySoRow = dtJson.NewRow();

                //        totalBySoRow["se_id"] = "";
                //        totalBySoRow["mer_po"] = "";
                //        totalBySoRow["prod_no"] = "";
                //        totalBySoRow["colorway"] = "";
                //        totalBySoRow["cr_reqdate"] = "";
                //        totalBySoRow["lpd"] = "";
                //        totalBySoRow["psdd"] = "";
                //        totalBySoRow["podd"] = "";
                //        totalBySoRow["mold_no"] = "Total";
                //        totalBySoRow["size_no"] = "";

                //        string[] qtyCols = {
                //    "se_qty", "cutQty", "stitchingQty", "assmeblyQty", "packingQty",
                //    "inStock_qty",
                //    "CuttingBalQty","StitchingBalQty","AssemblyBalQty","PackingBalQty","FGBalqty","ShipBalqty"
                //};


                //        foreach (var col in qtyCols)
                //            totalBySoRow[col] = rows.Sum(r => ToIntSafe(r[col]));

                //        int insertIndex = dtJson.Rows.IndexOf(rows.Last());
                //        dtJson.Rows.InsertAt(totalBySoRow, insertIndex + 1);
                //    }
                //    #endregion


                //    #region FINAL GRAND TOTAL ROW
                //    //var realRows = dtJson.AsEnumerable()
                //    //                     .Where(r => r.Field<string>("mold_no") != "Total" &&
                //    //                                 r.Field<string>("size_no") != "Total")
                //    //                     .ToList();
                //    var realRows = dtJson.AsEnumerable()
                //     .Where(r =>
                //         r.Field<string>("mold_no") != "Total" &&   // exclude Total_By_SO
                //         r.Field<string>("size_no") != "Total"     // exclude Final Total
                //     )
                //     .ToList();


                //    DataRow finalTotal = dtJson.NewRow();

                //    finalTotal["se_id"] = "";
                //    finalTotal["mer_po"] = "";
                //    finalTotal["prod_no"] = "";
                //    finalTotal["colorway"] = "";
                //    finalTotal["cr_reqdate"] = "";
                //    finalTotal["lpd"] = "";
                //    finalTotal["psdd"] = "";
                //    finalTotal["podd"] = "";
                //    finalTotal["mold_no"] = "";
                //    finalTotal["size_no"] = "Total";

                //    string[] allQtyCols = {
                //    "se_qty", "cutQty", "stitchingQty", "assmeblyQty", "packingQty",
                //    "inStock_qty",
                //    "CuttingBalQty","StitchingBalQty","AssemblyBalQty","PackingBalQty","FGBalqty","ShipBalqty"
                //};


                //    foreach (var col in allQtyCols)
                //        finalTotal[col] = realRows.Sum(r => ToIntSafe(r[col]));

                //    dtJson.Rows.Add(finalTotal);
                //    #endregion
                //    #region Bind to Grid

                //    dataGridView1.DataSource = dtJson.DefaultView;

                //    // Hide unused columns safely
                //    string[] columnsToHide =
                //    {
                //    "inStock_qty",
                //    "outSoleQty",
                //    "CRD_DUE_DAY_STATUS",
                //    "Column2",
                //    "lpd"
                //};

                //    if (dataGridView1 != null &&
                //        dataGridView1.DataSource != null &&
                //        dataGridView1.Columns.Count > 0)
                //    {
                //        foreach (string colName in columnsToHide)
                //        {
                //            if (dataGridView1.Columns.Contains(colName))
                //            {
                //                dataGridView1.Columns[colName].Visible = false;
                //            }
                //        }
                //    }

                //    ApplyDataGridViewStyles(dataGridView1);

                //    #endregion

                //    loadbl12.Visible = false;
                //    progressBar2.Visible = false;
                //    loadbl12.Visible = false;
                //    progressBar2.Visible = false;
                //}
                #endregion





                if (Convert.ToBoolean(Newtonsoft.Json.JsonConvert
    .DeserializeObject<Dictionary<string, object>>(ret)["IsSuccess"]))
                {
                    var dict = Newtonsoft.Json.JsonConvert.DeserializeObject<Dictionary<string, object>>(ret);

                    string json = dict["RetData"].ToString();
                    DataTable dtJson = SJeMES_Framework.Common.JsonHelper.GetDataTableByJson(json);

                    if (dtJson.Rows.Count == 0)
                    {
                        SJeMES_Control_Library.MessageHelper.ShowErr(this, "No such data");
                        loadbl12.Visible = false;
                        progressBar2.Visible = false;
                        return;
                    }

                    // 🔥 Run heavy logic in background
                    await Task.Run(() =>
                    {
                        ProcessData(dtJson);
                    });

                    // ==========================
                    // UI BIND (FAST)
                    // ==========================

                    dataGridView1.SuspendLayout();

                    dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
                    dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None;

                    typeof(DataGridView)
                        .GetProperty("DoubleBuffered", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)
                        .SetValue(dataGridView1, true, null);

                    dataGridView1.DataSource = dtJson.DefaultView;

                    dataGridView1.ResumeLayout();

                    // Hide columns
                    string[] columnsToHide =
                    {
                    "inStock_qty",
                    "outSoleQty",
                    "CRD_DUE_DAY_STATUS",
                    "Column2",
                    "lpd"
                };

                    foreach (string colName in columnsToHide)
                    {
                        if (dataGridView1.Columns.Contains(colName))
                            dataGridView1.Columns[colName].Visible = false;
                    }

                    ApplyDataGridViewStyles(dataGridView1);

                    loadbl12.Visible = false;
                    progressBar2.Visible = false;
                }
                else
                {
                    SJeMES_Control_Library.MessageHelper.ShowErr(this, Newtonsoft.Json.JsonConvert.DeserializeObject<Dictionary<string, object>>(ret)["ErrMsg"].ToString());
                    loadbl12.Visible = false;
                    progressBar2.Visible = false;
                }
            }

            if (tabControl1.SelectedIndex == 1) //Query Tab: Summary by Sales Order
            {


                Dictionary<string, Object> p = new Dictionary<string, object>();
                p.Add("vSeId", SOtxt.Text);
                p.Add("vPO", PO_txt.Text);
                p.Add("vCheckShip", checkBox6.Checked);
                p.Add("vCheckCRD", crdcheck.Checked);
                p.Add("vBeginDate", dateTimePicker5.Value.ToShortDateString());
                p.Add("vEndDate", dateTimePicker6.Value.ToShortDateString());
                p.Add("vCheckLPD", lpdcheck.Checked);
                p.Add("vLPDBeginDate", dateTimePicker7.Value.ToShortDateString());
                p.Add("vLPDEndDate", dateTimePicker8.Value.ToShortDateString());
                p.Add("plant", plantcm.SelectedItem);
                p.Add("SeIdList", bulktxt.Text.Trim().ToString());
                loadbl12.Visible = true;
                progressBar2.Style = ProgressBarStyle.Marquee;
                progressBar2.Visible = true;
                string ret = await Task.Run(() => SJeMES_Framework.WebAPI.WebAPIHelper.Post(Program.client.APIURL, "KZ_QCO", "KZ_QCO.Controllers.MESUpdateServer", "GetTotalDataBySeId_P", Program.client.UserToken, Newtonsoft.Json.JsonConvert.SerializeObject(p)));

                if (Convert.ToBoolean(Newtonsoft.Json.JsonConvert.DeserializeObject<Dictionary<string, object>>(ret)["IsSuccess"]))
                {


                    //   string json = Newtonsoft.Json.JsonConvert.DeserializeObject<Dictionary<string, object>>(ret)["RetData"].ToString();
                    //   DataTable dtJson = SJeMES_Framework.Common.JsonHelper.GetDataTableByJson(json);
                    ////   DataTable dtJson = JsonConvert.DeserializeObject<DataTable>(json);

                    //   if (dtJson.Rows.Count == 0 || dtJson == null)
                    //   {
                    //       SJeMES_Control_Library.MessageHelper.ShowErr(this, "No such data");
                    //       return;
                    //   }

                    string retDataJson = Newtonsoft.Json.JsonConvert
        .DeserializeObject<Dictionary<string, object>>(ret)["RetData"].ToString();

                    // Ensure RetData is an array
                    if (!retDataJson.TrimStart().StartsWith("["))
                        retDataJson = "[" + retDataJson + "]";

                     dt = ConvertJsonToDataTable(retDataJson);

                    if (dt == null || dt.Rows.Count == 0)
                    {
                        SJeMES_Control_Library.MessageHelper.ShowErr(this, "No such data");
                        return;
                    }


                    #region Add Balance Columns BEFORE Creating Total Row
                    void AddColumnIfMissing(string colName)
                    {
                        if (!dt.Columns.Contains(colName))
                            dt.Columns.Add(colName, typeof(int));
                    }
                    AddColumnIfMissing("CuttingBalQty2");
                    AddColumnIfMissing("StitchingBalQty2");
                    AddColumnIfMissing("AssemblyBalQty2");
                    AddColumnIfMissing("PackingBalQty2");
                    AddColumnIfMissing("FGBalqty2");
                    AddColumnIfMissing("ShipBalqty2");
                    #endregion

                    #region Calculate Balances
                    int ToIntSafe(object value)
                    {
                        return int.TryParse(Convert.ToString(value), out int result) ? result : 0;
                    }

                    foreach (DataRow row in dt.Rows)
                    {
                        if (row["mold_no"].ToString() == "Total") continue;

                        int order = ToIntSafe(row["se_qty"]);
                        int cut = ToIntSafe(row["cutQty"]);
                        int stitch = ToIntSafe(row["stitchingQty"]);
                        int assm = ToIntSafe(row["assmeblyQty"]);
                        int pack = ToIntSafe(row["packingQty"]);
                        int FG = ToIntSafe(row["inStock_qty"]);
                        int ship = ToIntSafe(row["shipping_qty"]);

                        row["CuttingBalQty2"] = Math.Max(order - cut, 0);
                        row["StitchingBalQty2"] = Math.Max(order - stitch, 0);
                        row["AssemblyBalQty2"] = Math.Max(order - assm, 0);
                        row["PackingBalQty2"] = Math.Max(order - pack, 0);
                        row["FGBalqty2"] = Math.Max(order - FG, 0);
                        row["ShipBalqty2"] = Math.Max(order - ship, 0);
                    }

                    #endregion
                    #region FINAL GRAND TOTAL ROW
                    var realRows = dt.AsEnumerable()
                     .Where(r => r.Field<string>("mold_no") != "Total")
                     .ToList();

                    DataRow totalRow = dt.NewRow();
                    totalRow["mold_no"] = "Total";

                    string[] qtyCols =
                    {
    "se_qty",
    "cutQty", "stitchingQty", "assmeblyQty", "packingQty", "inStock_qty",
    "CuttingBalQty2","StitchingBalQty2","AssemblyBalQty2",
    "PackingBalQty2","FGBalqty2","ShipBalqty2"
};

                    foreach (var col in qtyCols)
                    {
                        if (dt.Columns.Contains(col))
                            totalRow[col] = realRows.Sum(r => ToIntSafe(r[col]));
                    }

                    dt.Rows.Add(totalRow);

                    #endregion


                    #region This Block For Assembly Not Finished Then Show the assembly_finished_date Empty
                    if (dt.Columns.Contains("assembly_finished_date"))
                    {
                        foreach (DataRow row in dt.Rows)
                        {
                            // If not fully assembled → hide date
                            if (ToIntSafe(row["AssemblyBalQty2"]) != 0)
                            {
                                row["assembly_finished_date"] = "";
                            }

                        }
                    }
                    #endregion
                    dataGridView2.DataSource = dt.DefaultView;
                    dataGridView2.Columns["inStock_qty_2"].Visible = false;
                    dataGridView2.Columns["CRD_DUE_DAY_STATUS_SO"].Visible = false;

                    if (dataGridView2 != null &&
     dataGridView2.DataSource != null &&
     dataGridView2.Columns != null)
                    {
                        int[] hideIndexes = { 3, 4, 6, 8 };

                        foreach (int idx in hideIndexes)
                        {
                            if (dataGridView2.Columns.Count > idx)
                            {
                                dataGridView2.Columns[idx].Visible = false;
                            }
                        }
                    }
                    //    dataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                    ApplyDataGridViewStyles(dataGridView2);
                    loadbl12.Visible = false;
                    progressBar2.Visible = false;
                }
                else
                {
                    SJeMES_Control_Library.MessageHelper.ShowErr(this, Newtonsoft.Json.JsonConvert.DeserializeObject<Dictionary<string, object>>(ret)["ErrMsg"].ToString());
                    loadbl12.Visible = false;
                    progressBar2.Visible = false;
                }

            }
        }
        void ProcessData(DataTable dtJson)
        {
            // ======================
            // ADD COLUMNS (SAFE)
            // ======================
            void AddColumn(string col)
            {
                if (!dtJson.Columns.Contains(col))
                    dtJson.Columns.Add(col, typeof(int));
            }

            AddColumn("CuttingBalQty");
            AddColumn("StitchingBalQty");
            AddColumn("AssemblyBalQty");
            AddColumn("PackingBalQty");
            AddColumn("FGBalqty");
            AddColumn("ShipBalqty");

            // ======================
            // HELPER
            // ======================
            int ToIntSafe(object value)
            {
                if (value == null || value == DBNull.Value) return 0;

                if (int.TryParse(value.ToString(), out int i))
                    return i;

                if (decimal.TryParse(value.ToString(), out decimal d))
                    return (int)d;

                return 0;
            }

            bool IsFinishedLine(DataRow r)
            {
                bool isSupplier = ToIntSafe(r["IS_SUPPLIER_ORDER"]) == 1;

                if (isSupplier)
                {
                    return ToIntSafe(r["AssemblyBalQty"]) <= 0 &&
                           ToIntSafe(r["PackingBalQty"]) <= 0;
                }
                else
                {
                    return ToIntSafe(r["CuttingBalQty"]) <= 0 &&
                           ToIntSafe(r["StitchingBalQty"]) <= 0 &&
                           ToIntSafe(r["AssemblyBalQty"]) <= 0 &&
                           ToIntSafe(r["PackingBalQty"]) <= 0;
                }
            }

            string FormatDate(object val)
            {
                if (val == DBNull.Value || string.IsNullOrWhiteSpace(val?.ToString()))
                    return "";

                if (DateTime.TryParse(val.ToString(), out DateTime d))
                    return d.ToString("yyyy-MM-dd");

                return "";
            }

            bool hasAssemblyDate = dtJson.Columns.Contains("assembly_finished_date");
            bool hasPCDate = dtJson.Columns.Contains("PC_DATE");

            // ======================
            // DATE FORMAT (SAME)
            // ======================
            if (hasAssemblyDate)
            {
                foreach (DataRow row in dtJson.Rows)
                    row["assembly_finished_date"] = FormatDate(row["assembly_finished_date"]);

                dtJson.Columns["assembly_finished_date"].DataType = typeof(string);
            }

            if (hasPCDate)
            {
                foreach (DataRow row in dtJson.Rows)
                    row["PC_DATE"] = FormatDate(row["PC_DATE"]);

                dtJson.Columns["PC_DATE"].DataType = typeof(string);
            }

            // ======================
            // BALANCE CALCULATION
            // ======================
            foreach (DataRow row in dtJson.Rows)
            {
                int order = ToIntSafe(row["se_qty"]);
                int cut = ToIntSafe(row["cutQty"]);
                int stitch = ToIntSafe(row["stitchingQty"]);
                int assm = ToIntSafe(row["assmeblyQty"]);
                int pack = ToIntSafe(row["packingQty"]);
                int FG = ToIntSafe(row["inStock_qty"]);
                int ship = ToIntSafe(row["shipping_qty"]);

                row["CuttingBalQty"] = Math.Max(order - cut, 0);
                row["StitchingBalQty"] = Math.Max(order - stitch, 0);
                row["AssemblyBalQty"] = Math.Max(order - assm, 0);
                row["PackingBalQty"] = Math.Max(order - pack, 0);
                row["FGBalqty"] = Math.Max(order - FG, 0);
                row["ShipBalqty"] = Math.Max(order - ship, 0);
            }

            // ======================
            // HIDE DATE LOGIC (SAME)
            // ======================
            if (hasAssemblyDate)
            {
                foreach (DataRow row in dtJson.Rows)
                {
                    if (ToIntSafe(row["PackingBalQty"]) != 0)
                        row["assembly_finished_date"] = "";
                }
            }

            // ======================
            // FILTER (FASTER)
            // ======================
            if (checkBox3.Checked)
            {
                for (int i = dtJson.Rows.Count - 1; i >= 0; i--)
                {
                    var r = dtJson.Rows[i];

                    if (!string.IsNullOrWhiteSpace(r["size_no"].ToString()) &&
                        r["mold_no"].ToString() != "Total" &&
                        IsFinishedLine(r))
                    {
                        dtJson.Rows.RemoveAt(i);
                    }
                }

                dtJson.AcceptChanges();
            }

            // ======================
            // GROUP TOTAL (ORIGINAL STYLE)
            // ======================
            var seIds = dtJson.AsEnumerable()
                .Select(r => r.Field<string>("se_id"))
                .Distinct()
                .ToList();

            string[] qtyCols = {
        "se_qty","cutQty","stitchingQty","assmeblyQty","packingQty",
        "inStock_qty",
        "CuttingBalQty","StitchingBalQty","AssemblyBalQty",
        "PackingBalQty","FGBalqty","ShipBalqty"
    };

            foreach (var seId in seIds)
            {
                var rows = dtJson.AsEnumerable()
                    .Where(r => r.Field<string>("se_id") == seId &&
                                r.Field<string>("size_no") != "Total_By_SO" &&
                                r.Field<string>("size_no") != "Total")
                    .ToList();

                if (!rows.Any()) continue;

                DataRow totalRow = dtJson.NewRow();

                totalRow["mold_no"] = "Total";
                totalRow["size_no"] = "";

                foreach (var col in qtyCols)
                    totalRow[col] = rows.Sum(r => ToIntSafe(r[col]));

                int index = dtJson.Rows.IndexOf(rows.Last());
                dtJson.Rows.InsertAt(totalRow, index + 1);
            }

            // ======================
            // FINAL TOTAL (SAME)
            // ======================
            var realRows = dtJson.AsEnumerable()
                .Where(r => r.Field<string>("mold_no") != "Total" &&
                            r.Field<string>("size_no") != "Total")
                .ToList();

            DataRow finalTotal = dtJson.NewRow();
            finalTotal["size_no"] = "Total";

            foreach (var col in qtyCols)
                finalTotal[col] = realRows.Sum(r => ToIntSafe(r[col]));

            dtJson.Rows.Add(finalTotal);
        }

        public static DataTable ConvertJsonToDataTable(string json)
        {
            var array = Newtonsoft.Json.Linq.JArray.Parse(json);
            DataTable dt = new DataTable();

            // Add columns dynamically
            foreach (var prop in array.First.Children<Newtonsoft.Json.Linq.JProperty>())
                dt.Columns.Add(prop.Name);

            // Add rows
            foreach (var obj in array.Children<Newtonsoft.Json.Linq.JObject>())
            {
                DataRow dr = dt.NewRow();
                foreach (var prop in obj.Properties())
                    dr[prop.Name] = prop.Value?.ToString();
                dt.Rows.Add(dr);
            }

            return dt;
        }


        private bool Validate_CRD_Date()
        {
            DateTime fromDate = dateTimePicker5.Value.Date;
            DateTime toDate = dateTimePicker6.Value.Date;

            DateTime maxAllowedDate = fromDate.AddMonths(3); // strictly 3 calendar months

            if (toDate > maxAllowedDate)
            {
                return false;
            }
            return true;
        }

        private bool Validate_LPD_Date()
        {
            DateTime fromDate = dateTimePicker7.Value.Date;
            DateTime toDate = dateTimePicker8.Value.Date;

            DateTime maxAllowedDate = fromDate.AddMonths(3); // strictly 3 calendar months

            if (toDate > maxAllowedDate)
            {
                return false;
            }
            return true;
        }
        private Dictionary<string, DataTable> GetAssmeblyQtyDetail(string se_id, string size_no, string vProcessNo)
        {
            Dictionary<string, Object> p = new Dictionary<string, object>();

            p.Add("vSeId", se_id);
            p.Add("vSizeNo", size_no);
            p.Add("vProcessNo", vProcessNo);
            string ret = SJeMES_Framework.WebAPI.WebAPIHelper.Post(Program.client.APIURL, "KZ_QCO", "KZ_QCO.Controllers.MESUpdateServer", "CusGetFinishQtyDetail", Program.client.UserToken, Newtonsoft.Json.JsonConvert.SerializeObject(p));
            if (Convert.ToBoolean(Newtonsoft.Json.JsonConvert.DeserializeObject<Dictionary<string, object>>(ret)["IsSuccess"]))
            {
                string json = Newtonsoft.Json.JsonConvert.DeserializeObject<Dictionary<string, object>>(ret)["RetData"].ToString();
                Dictionary<string, DataTable> jarr = Newtonsoft.Json.JsonConvert.DeserializeObject<Dictionary<string, DataTable>>(json);

                return jarr;
            }
            else
            {
                SJeMES_Control_Library.MessageHelper.ShowErr(this, Newtonsoft.Json.JsonConvert.DeserializeObject<Dictionary<string, object>>(ret)["ErrMsg"].ToString());
                return null;
            }
        }
        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            int ColumnIndex = e.ColumnIndex;


            if (CusFormIndex.ContainsKey(ColumnIndex))  //Open the [Number of Work Reports] viewing interface
            {

                DataGridView dgv = (DataGridView)sender;
                string se_id = dgv.Rows[e.RowIndex].Cells["se_id"].Value.ToString();
                string size_no = dgv.Rows[e.RowIndex].Cells["size_no"].Value.ToString();
                if (string.IsNullOrEmpty(se_id) || string.IsNullOrEmpty(size_no))
                {
                    return;
                }
                CusAssmeblyQtyForm cusAssmebly = new CusAssmeblyQtyForm(dgv.Columns[ColumnIndex].HeaderText, GetAssmeblyQtyDetail(se_id, size_no, CusFormIndex[ColumnIndex]));
                cusAssmebly.ShowDialog();

            }



            if (ColumnIndex == dataGridView1.Columns["outSoleQty"].Index)  //Open the [Background Inventory] data viewing interface
            {
                DataGridView dgv = (DataGridView)sender;
                OutSoleQtyForm frm = new OutSoleQtyForm(dgv.Rows[e.RowIndex].Cells["se_id"].Value.ToString(), dgv.Rows[e.RowIndex].Cells["size_no"].Value.ToString());
                frm.StartPosition = FormStartPosition.CenterParent;
                frm.ShowDialog();
            }


            if (ColumnIndex == dataGridView1.Columns["inStock_qty"].Index)  //Open the [Inbound Quantity] viewing interface
            {
                DataGridView dgv = (DataGridView)sender;
                InStockQtyForm frm = new InStockQtyForm(dgv.Rows[e.RowIndex].Cells["se_id"].Value.ToString(), dgv.Rows[e.RowIndex].Cells["size_no"].Value.ToString());
                frm.StartPosition = FormStartPosition.CenterParent;
                frm.ShowDialog();
            }

            if (ColumnIndex == dataGridView1.Columns["shipping_qty"].Index)  //Open the [Shipment Quantity] view interface
            {
                DataGridView dgv = (DataGridView)sender;
                ShippingQtyForm frm = new ShippingQtyForm(dgv.Rows[e.RowIndex].Cells["se_id"].Value.ToString(), dgv.Rows[e.RowIndex].Cells["size_no"].Value.ToString());
                frm.StartPosition = FormStartPosition.CenterParent;
                frm.ShowDialog();
            }
        }



     
        private void ExportExcels(string fileName, DataGridView myDGV)
        {
            SaveFileDialog saveDialog = new SaveFileDialog();
            saveDialog.DefaultExt = "xlsx";
            saveDialog.Filter = "Excel文件|*.xlsx";
            saveDialog.FileName = fileName;
            if (saveDialog.ShowDialog() != DialogResult.OK) return;

            string saveFileName = saveDialog.FileName;

            var xlApp = new Microsoft.Office.Interop.Excel.Application();
            var workbook = xlApp.Workbooks.Add();
            var worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets[1];

            // ➤ Columns you want to skip in Export
            string[] skipColumns = { "outSoleQty", /*"shipping_qty",*/ "CRD_DUE_DAY_STATUS_SO", "CRD_DUE_DAY_STATUS", "Column2" , "dataGridViewTextBoxColumn5" };

            int colIndex = 1;

            // -------- 1️⃣ Write Header --------
            foreach (DataGridViewColumn col in myDGV.Columns)
            {
                if (skipColumns.Contains(col.Name))
                    continue;    // 🔥 Skip ONLY selected columns

                worksheet.Cells[1, colIndex] = "'" + col.HeaderText;
                colIndex++;
            }

            // -------- 2️⃣ Write Data --------
            int exportColumnCount = myDGV.Columns
                .Cast<DataGridViewColumn>()
                .Count(c => !skipColumns.Contains(c.Name));

            object[,] objData = new object[myDGV.Rows.Count, exportColumnCount];

            for (int i = 0; i < myDGV.Rows.Count; i++)
            {
                colIndex = 0;

                foreach (DataGridViewColumn col in myDGV.Columns)
                {
                    if (skipColumns.Contains(col.Name))
                        continue;  // 🔥 Skip selected

                    objData[i, colIndex] = myDGV.Rows[i].Cells[col.Index].FormattedValue.ToString();
                    colIndex++;
                }
            }

            var rg = worksheet.Range[
                worksheet.Cells[2, 1],
                worksheet.Cells[myDGV.Rows.Count + 1, exportColumnCount]
            ];
            rg.Value2 = objData;

            worksheet.Columns.AutoFit();
            workbook.SaveAs(saveFileName);
            workbook.Close();
            xlApp.Quit();
            GC.Collect();

            MessageBox.Show("Export Successful!\nడేటా ఎగుమతి విజయవంతంగా పూర్తయింది!");

        }


        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 2)
            {
                textBox_DeptNo.Enabled = true;
            }
            else
            {
                textBox_DeptNo.Enabled = false;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            PO_txt.Text = "";
            SOtxt.Text = "";
            // Item_no_List.Text = "";
            bulktxt.Text = "";
        }

        private void RichTextBox1_TextChanged(object sender, EventArgs e)
        {
            if (!(sender is RichTextBox)) { return; }
            RichTextBox richText = sender as RichTextBox;
            if (richText.Text.Contains("\n"))
            {
                ExcelFormat(richText);
            }
        }
        private void ExcelFormat(RichTextBox richText)
        {
            string[] str = richText.Text.Split(new string[] { "\t\n" }, StringSplitOptions.RemoveEmptyEntries);
            if (str.Length <= 1)
                str = richText.Text.Split(new string[] { "\n" }, StringSplitOptions.RemoveEmptyEntries);

            string se_id = "";
            for (int i = 0; i < str.Length; i++)
            {
                if (se_id.Length > 0)
                {
                    se_id += ",";
                }
                se_id += str[i];

            }
            richText.Text = se_id;
            richText.Font = new System.Drawing.Font("宋体", 9F);

        }

        private void dataGridView1_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            var row = dataGridView1.Rows[e.RowIndex];

            if (!row.Cells.Contains(row.Cells["CRD_DUE_DAY_STATUS"]))
                return;

            object val = row.Cells["CRD_DUE_DAY_STATUS"].Value;

            if (val == null || val == DBNull.Value)
                return;

            if (!int.TryParse(val.ToString().Trim(), out int crdStatus))
                return;

            string[] targetCols =
            {
        "CuttingBalQty",
        "StitchingBalQty",
        "AssemblyBalQty",
        "PackingBalQty",
        "FGBalqty",
        "shipping_qty"
    };

            Color backColor = crdStatus == 1 ? Color.LightCoral : Color.PaleGreen;

            foreach (string col in targetCols)
            {
                if (row.Cells[col] != null)
                    row.Cells[col].Style.BackColor = backColor;
            }
        }

        private void dataGridView2_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            var row = dataGridView2.Rows[e.RowIndex];

            if (!row.Cells.Contains(row.Cells["CRD_DUE_DAY_STATUS_SO"]))
                return;

            object val = row.Cells["CRD_DUE_DAY_STATUS_SO"].Value;

            if (val == null || val == DBNull.Value)
                return;

            if (!int.TryParse(val.ToString().Trim(), out int crdStatus))
                return;

            string[] targetCols =
            {
        "CuttingBalQty2",
        "StitchingBalQty2",
        "AssemblyBalQty2",
        "PackingBalQty2",
        "FGBalqty2",
        "shipping_qty_2"
    };

            Color backColor = crdStatus == 1 ? Color.LightCoral : Color.PaleGreen;

            foreach (string col in targetCols)
            {
                if (row.Cells[col] != null)
                    row.Cells[col].Style.BackColor = backColor;
            }

        }
       
        private void ApplyDataGridViewStyles(DataGridView dgv)
        {
            if (dgv == null) return;

            dgv.EnableHeadersVisualStyles = false;

            dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.Teal;
            dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgv.ColumnHeadersDefaultCellStyle.Font = new Font("Times New Roman", 10, FontStyle.Bold);

            dgv.DefaultCellStyle.Font = new Font("Times New Roman", 11, FontStyle.Regular);
            dgv.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            dgv.GridColor = Color.Teal;

            // 🔹 This makes columns fit inside the DataGridView width
            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

            // Optional: remove horizontal scroll
            dgv.ScrollBars = ScrollBars.Vertical;
        }

       

        private void Expor_Click(object sender, EventArgs e)
        {
            string fileName = "PO Tracking.xls";

            if (tabControl1.SelectedTab == tabPage1)
            {
                if (dataGridView1.Rows.Count > 0)
                    ExportExcels(fileName, dataGridView1);
                else
                    MessageBox.Show("No data in first grid.");
            }
            else if (tabControl1.SelectedTab == tabPage2)
            {
                if (dataGridView2.Rows.Count > 0)
                    ExportExcels(fileName, dataGridView2);
                else
                    MessageBox.Show("No data in second grid.");
            }
        }

       
    }
}

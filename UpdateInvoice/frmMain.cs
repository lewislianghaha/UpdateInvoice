using System;
using System.Data;
using System.Windows.Forms;
using UpdateInvoice.Interface;

namespace UpdateInvoice
{
    public partial class FrmMain : Form
    {
        public FrmMain()
        {
            InitializeComponent();
            OnRegisterEvents();
        }

        private void OnRegisterEvents()
        {
            btnOpenExcel.Click += btnOpenExcel_Click;
            btnImport.Click += btnImport_Click;
            btnExit.Click += btnExit_Click;
        }


        /// <summary>
        /// 打开EXCEL
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void btnOpenExcel_Click(object sender, EventArgs e)
        {
            var openFileDialog = new OpenFileDialog { Filter = "Xlsx文件|*.xlsx" };
            if (openFileDialog.ShowDialog() != DialogResult.OK) return;
            try
            {
                string strPath = openFileDialog.FileName;
                var exc = new Excel();
                var dt = exc.OpenExcel(strPath);
                gvdtl.DataSource = dt;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }


        /// <summary>
        /// 导入功能
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void btnImport_Click(object sender, EventArgs e)
        {
            var dt = new DataTable();
            var exec = new Excel();
            int clickid;          //记录单选项标记
            string clickMessage; //记录选择来源
            var result = new FrmErrorResult();

            try
            {
                if (gvdtl.Rows.Count == 0) throw new Exception("没有EXCEL内容,请导入后再继续");

                clickid = rdSale.Checked ? 0 : 1;
                clickMessage = clickid == 0 ? "您所选择导入的单据来源为销售发票,是否继续?" : "您所选择导入的单据来源为采购发票,是否继续?";

                if (MessageBox.Show(clickMessage, "提示", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                {
                    //将DataGridView控件内的数据转变为DataTable
                    dt = (DataTable)gvdtl.DataSource;

                    //获取验证从DataGridView里的DataTable能更新的值
                    var canImportdt = exec.VaildCanImportTable(dt, clickid);
                    //获取验证从DataGridView里的DataTable能更新的值
                    var cannotImportdt = exec.ErrorTable;

                    //若验证的结果与DataGridView一致,即表示全部能更新
                    if (dt.Rows.Count == canImportdt.Rows.Count)
                    {
                        exec.ImportExcel(canImportdt,clickid);
                        MessageBox.Show("已成功更新,请到K3系统对应的单据进行查阅", "提示信息", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                    //若不能更新的记录与原DataTable一致的话
                    else if (cannotImportdt.Rows.Count==dt.Rows.Count)
                    {
                        var errormessage = "抱歉地通知您,Excel里所有记录都不能成功导入\n 请整理数据后再进行更新";
                        MessageBox.Show(errormessage, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    else if (canImportdt.Rows.Count > cannotImportdt.Rows.Count || canImportdt.Rows.Count == cannotImportdt.Rows.Count || canImportdt.Rows.Count < cannotImportdt.Rows.Count)
                    {

                        exec.ImportExcel(canImportdt,clickid);
                        var message = "已成为更新一部份信息,不能更新的信息请按确定进行查阅";
                        MessageBox.Show(message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        result.LoadErrorRecord(cannotImportdt);
                        result.ShowDialog();
                    }
                }

                //清空原来DataGridView内的内容(无论成功与否都会执行)
                var dt1 = (DataTable)gvdtl.DataSource;
                dt1.Rows.Clear();
                gvdtl.DataSource = dt1;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message,"错误",MessageBoxButtons.OK,MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 退出功能
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void btnExit_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("是否退出?", "提示", MessageBoxButtons.YesNo,MessageBoxIcon.Information) == DialogResult.Yes)
                Close();
        }
    }
}

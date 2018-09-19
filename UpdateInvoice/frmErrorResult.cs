using System;
using System.Collections;
using System.Data;
using System.Windows.Forms;
using UpdateInvoice.Interface;

namespace UpdateInvoice
{
    public partial class FrmErrorResult : Form
    {
        //private DataTable _errorTabledt;

        //public DataTable DtTable
        //{
        //    //接收错误表格信息
        //    set { _errorTabledt = value; }
        //}

        public FrmErrorResult()
        {
            InitializeComponent();
            OnRegisterEvents();
            // LoadErrorRecord();
        }

        private void OnRegisterEvents()
        {
            btnexport.Click += btnexport_Click;
            btnExit.Click += btnExit_Click;
        }

        //获取错误结果集并绑定至DataGridView中
        public void LoadErrorRecord(DataTable dt)
        {
            gverrordtl.DataSource = dt;
            txtrows.Text = Convert.ToString(gverrordtl.Rows.Count);
        }

        /// <summary>
        /// 导出EXCEL
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void btnexport_Click(object sender, EventArgs e)
        {
            var excel = new Excel();

            try
            {
                if (MessageBox.Show("是否将错误信息导出至Excel", "提示", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                {

                    var saveFileDialog = new SaveFileDialog { Filter = "Xlsx文件|*.xlsx" };
                    if (saveFileDialog.ShowDialog() != DialogResult.OK) return;
                    var result = excel.ExportExcel(saveFileDialog.FileName, (DataTable)gverrordtl.DataSource);

                    if (result["Code"].ToString() == "0")
                    {
                        MessageBox.Show("导出成功,请到原导出的位置找到Excel文件进行查阅", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        Close();
                    }
                    else
                    {
                        MessageBox.Show(result["Code"].ToString(), "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 退出
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void btnExit_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("是否退出?", "提示", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                Close();
        }
    }
}

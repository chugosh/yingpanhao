using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using CCWin;
using System.IO;

namespace WindowsFormsApplication1
{
    public partial class Form1 : CCSkinMain
    {
        int work = 2101;
        VibrationTest vt = new VibrationTest();
        DrillingTest dt = new DrillingTest();
        
        public Form1()
        { 
            InitializeComponent();
        }
        
        private void Form1_Load(object sender, EventArgs e)
        {

        }
        
        //钻屑法按钮
        private void skinButton1_Click(object sender, EventArgs e)
        {
            dt.Start();
            MessageBox.Show("ok");
        }
        //平面图按钮
        private void skinButton2_Click(object sender, EventArgs e)
        {
            string filePath = null;
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "选择要上传的图片";
            //图片格式
            openFileDialog.Filter = "*.png|*.png|*.jpg|*.jpg|*.bmp|*.bmp";
            //不允许多选
            openFileDialog.Multiselect = false;
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                filePath = openFileDialog.FileName;
                skinTextBox1.Text = filePath;
                skinPictureBox1.Load(filePath);
            }
            vt.AddPictures(filePath,null);
            MessageBox.Show("picture ok");
        }
        //剖面图按钮
        private void skinButton4_Click(object sender, EventArgs e)
        {
            string filePath = null;
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "选择要上传的图片";
            //图片格式
            openFileDialog.Filter = "*.png|*.png|*.jpg|*.jpg|*.bmp|*.bmp";
            //不允许多选
            openFileDialog.Multiselect = false;
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                filePath = openFileDialog.FileName;
                skinTextBox2.Text = filePath;
                skinPictureBox2.Load(filePath);
            }
            vt.AddPictures(null, filePath);
            MessageBox.Show("picture2 ok");
        }
        //微震按钮
        private void skinButton3_Click(object sender, EventArgs e)
        {
            vt.GetDataTable();
            MessageBox.Show("微震监测ok");
        }
        //应力监测按钮
        private void skinButton5_Click(object sender, EventArgs e)
        {
            StressTest st = new StressTest(work);
            MessageBox.Show("应力监测ok");
        }


/*----------------------------------------------------------------------------------*/

        private void SavePicuture()
        {
            //保存图片
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.ShowDialog();
            skinPictureBox1.Image.Save(saveFileDialog.FileName);
        }
        private void skinTextBox1_Paint(object sender, PaintEventArgs e)
        {

        }
        private void skinTextBox2_Paint(object sender, PaintEventArgs e)
        {

        }
    }
}

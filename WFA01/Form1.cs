using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WFA01
{
    public partial class Form1 : Form
    {
        public bool isProgress = false;
        public int processPercentage = 0;

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog()
            {
                Multiselect = false,  // 複数選択の可否
                Filter =  // フィルタ
                "Excelファイル|*.xlsx;",
            };

            //ダイアログを表示
            Stream myStream = null;
            DialogResult result = dialog.ShowDialog();
            if (result == DialogResult.OK)
            {
                // ダイアログで指定されたファイルを読み込み
                try
                {
                    if ((myStream = dialog.OpenFile()) != null)
                    {
                        // ファイル名を設定
                        textBox1.Text = dialog.FileName;
                        Converter.file_path = dialog.FileName;
                        // 実行ボタンを有効に
                        button2.Enabled = true;
                        myStream.Close();
                    }
                }
                catch (Exception e1)
                {
                    MessageBox.Show(e1.Message + "\n\n指定したファイルを Excel などで開いている場合は、いったん閉じてください。", "ファイル読み込みエラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private async void button2_Click(object sender, EventArgs e)
        {
            if (isProgress)
            {
                setButtonText("データを取得");
                addLog("中止しました");
                processProgress(0);
                isProgress = false;
            }
            else
            {
                var p = new Progress<int>(processProgress);
                var t = new Progress<string>(addLog);
                isProgress = true;
                setButtonText("中止する");
                await Task.Run(() => Converter.Convert(p, t));
                //Converter.Convert();
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Converter.checkbox1_state = checkBox1.Checked;
            Converter.checkbox2_state = checkBox2.Checked;
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            Converter.checkbox1_state = checkBox1.Checked;
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            Converter.checkbox2_state = checkBox2.Checked;
        }

        public void setButtonText(string s)
        {
            button2.Text = s;
        }

        public void addLog(string s)
        {
            textBox2.AppendText(s + "\n");
        }

        public void setProgressMax(int m)
        {
            progressBar1.Maximum = m;
        }

        private void processProgress(int percent)
        {
            progressBar1.Value = percent;
            processPercentage = percent;
            if (percent == 100)
            {
                setButtonText("データを取得");
                progressBar1.Value = 0;
                isProgress = false;
            }
        }

        public void resetProgress()
        {
            progressBar1.Value = 0;
        }

    }
}

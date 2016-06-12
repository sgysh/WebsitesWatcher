using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OpenCvSharp;

namespace WebsitesWatcher
{
    public struct ReceiveData
    {
        public string url;
        public bool is_canceled;
    }

    public partial class Form1 : Form
    {
        const int DVASPECT_CONTENT = 1;
        [System.Runtime.InteropServices.DllImport("ole32.dll")]
        extern static int OleDraw(
            IntPtr pUnk,
            int dwAspect,
            IntPtr hdcDraw,
            ref Rectangle lprcBounds);

        private bool IsFirstTime = true;
        private bool isChecking = false;

        public static readonly string FILE_PATH = Application.UserAppDataPath + "\\";
        public static readonly string XML_PATH = FILE_PATH + "test_datamanagement.xml";
        public DataTable dataTable;
        public static readonly string TABLE_NAME = "url_table";
        public static readonly string CHECK_COLUMN_NAME = "Check";
        public static readonly string URL_COLUMN_NAME = "URL";
        public static readonly string VALUE_COLUMN_NAME = "VALUE";
        public static readonly string DATE_COLUMN_NAME = "DATE";
        public static readonly string CHANGED_COLUMN_NAME = "CHANGED";
        public List<string> urls = new List<string>();
        int list_index = 0;
        public static readonly string OLD_FILE_SUFFIX = "_old.jpg";
        public static readonly string NEW_FILE_SUFFIX = "_new.jpg";

        enum SimilariryType { Initialized, NotChange, ALittleChange, Change };

        public Form1()
        {
            InitializeComponent();
        }

        private void ShowMessageOnBalloon(string str, int sec)
        {
            notifyIcon1.BalloonTipText = str;
            notifyIcon1.ShowBalloonTip(1000 * sec);
        }

        private void StartCheck()
        {

            list_index = 0;
            urls.Clear();
            foreach (DataRow item in dataSet1.Tables[TABLE_NAME].Rows)
            {
                urls.Add(item[URL_COLUMN_NAME].ToString());
            }
            webBrowser1.Navigate(urls[list_index]);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            MaximumSize = new Size(this.Width, Screen.PrimaryScreen.Bounds.Height);
            MinimumSize = new Size(this.Width, 0);

            webBrowser1.Visible = false;
            pictureBox1.Visible = false;
            // WebBrowserのサイズを変更可能にする
            webBrowser1.Dock = DockStyle.None;
            // WebBrowserのスクロールバーを表示しない
            webBrowser1.ScrollBarsEnabled = false;
            // ダイアログボックスを表示しない
            webBrowser1.ScriptErrorsSuppressed = true;

            if (System.IO.File.Exists(XML_PATH))
            {
                dataSet1.ReadXml(XML_PATH, XmlReadMode.InferTypedSchema);
                dataTable = dataSet1.Tables[TABLE_NAME];
            }
            else {
                dataTable = new DataTable(TABLE_NAME);
                dataTable.Columns.Add(new DataColumn(CHECK_COLUMN_NAME, typeof(bool)));
                dataTable.Columns.Add(new DataColumn(URL_COLUMN_NAME, typeof(string)));
                dataTable.Columns.Add(new DataColumn(VALUE_COLUMN_NAME, typeof(string)));
                dataTable.Columns.Add(new DataColumn(DATE_COLUMN_NAME, typeof(string)));
                dataTable.Columns.Add(new DataColumn(CHANGED_COLUMN_NAME, typeof(bool)));
                dataSet1.Tables.Add(dataTable);
            }

            dataGridView1.DataSource = dataSet1;
            dataGridView1.DataMember = TABLE_NAME;

            this.dataGridView1.RowHeadersVisible = false;
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToResizeColumns = false;
            this.dataGridView1.AllowUserToResizeRows = false;
            this.dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            this.dataGridView1.Columns[dataGridView1.ColumnCount - 1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            this.dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            this.dataGridView1.Columns[1].ReadOnly = true;
        }

        private string GetHash(string str)
        {
            var buf = System.Text.Encoding.UTF8.GetBytes(str);
            System.Security.Cryptography.MD5 md5 = System.Security.Cryptography.MD5.Create();

            var hash = md5.ComputeHash(buf);
            md5.Clear();

            return BitConverter.ToString(hash).ToLower().Replace("-", "");
        }

        private void webBrowser1_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            if (e.Url != ((WebBrowser)sender).Url) return;

            if (IsFirstTime)
            {
                IsFirstTime = false;
                webBrowser1.Navigate(webBrowser1.Url);
                return;
            }

            // workaround
            System.Threading.Thread.Sleep(300);
            Application.DoEvents();

            // ページを全体表示(スクロールバー無)するため、サイズをページの本文サイズにあわせる
            webBrowser1.Width  = webBrowser1.Document.Body.ScrollRectangle.Width;
            webBrowser1.Height = webBrowser1.Document.Body.ScrollRectangle.Height;

            // workaround
            System.Threading.Thread.Sleep(300);
            Application.DoEvents();

            TaskAsnc();
        }

        private async void TaskAsnc()
        {
            progressBar1.Step = progressBar1.Maximum / dataSet1.Tables[TABLE_NAME].Rows.Count;

            double distance = -1;
            await Task.Run(() =>
            {
                distance = CompareWebPage();
            });

            UpdateCurrentRow(distance);

            progressBar1.PerformStep();
            ++list_index;

            if (list_index < urls.Count)
            {
                webBrowser1.Navigate(urls[list_index]);
            }
            if (list_index == urls.Count)
            {
                isChecking = false;
                progressBar1.Value = 0;
            }
        }

        private double CompareWebPage()
        {
            // workaround
            System.Threading.Thread.Sleep(3000);
            Application.DoEvents();

            // WebBrowserのサイズに合わせてBitmap生成
            Bitmap bmp = new Bitmap(webBrowser1.Width, webBrowser1.Height);

            // BitmapのGraphicsを取得
            Graphics gra = Graphics.FromImage(bmp);

            // BitmapのGraphicsのHdcを取得
            IntPtr hdc = gra.GetHdc();

            // workaround
            System.Threading.Thread.Sleep(3000);
            Application.DoEvents();

            // WebBrowserのオブジェクト取得
            IntPtr web =
                System.Runtime.InteropServices.Marshal.GetIUnknownForObject(
                webBrowser1.ActiveXInstance);

            // workaround
            System.Threading.Thread.Sleep(3000);
            Application.DoEvents();

            // WebBrowserのイメージをBitmapにコピー
            Rectangle rect = new Rectangle(0, 0, bmp.Width, bmp.Height);
            OleDraw(web, DVASPECT_CONTENT, hdc, ref rect);

            // WebBrowserのオブジェクト使用終了
            System.Runtime.InteropServices.Marshal.Release(web);

            // BitmapのGraphicsの使用終了
            gra.Dispose();

            var name = GetHash(webBrowser1.Url.ToString());
            var new_filename = FILE_PATH + name + NEW_FILE_SUFFIX;
            var old_filename = FILE_PATH + name + OLD_FILE_SUFFIX;

            if (System.IO.File.Exists(old_filename))
            {
                System.IO.File.Delete(old_filename);
            }
            if (System.IO.File.Exists(new_filename))
            {
                System.IO.File.Move(new_filename, old_filename);
            }

            // BitmapをJPEGで保存
            bmp.Save(new_filename, System.Drawing.Imaging.ImageFormat.Jpeg);

            double distance = 0;
            if (!System.IO.File.Exists(old_filename))
            {
                distance = -1;
            }
            else

            {
                distance = CalculateHistogramDistance(new_filename, old_filename);
            }

            return distance;
        }

        private SimilariryType ToSimilariryType(double distance)
        {
            SimilariryType type;

            if (distance == -1)
            {
                type = SimilariryType.Initialized;
            }
            else if (distance < 0.02)
            {
                type = SimilariryType.NotChange;
            }
            else if (distance < 0.04)
            {
                type = SimilariryType.ALittleChange;
            }
            else
            {
                type = SimilariryType.Change;
            }

            return type;
        }

        private void UpdateCurrentRow(double distance)
        {
            string value_str;
            var type = ToSimilariryType(distance);

            foreach (DataRow row in dataSet1.Tables[TABLE_NAME].Rows)
            {
                if (urls[list_index] == row[URL_COLUMN_NAME].ToString())
                {
                    switch (type)
                    {
                        case SimilariryType.Initialized:
                            value_str = "Initialized";
                            break;
                        case SimilariryType.NotChange:
                            value_str = "Not change";
                            dataGridView1[VALUE_COLUMN_NAME, list_index].Style.BackColor = Color.Green;
                            break;
                        case SimilariryType.ALittleChange:
                            value_str = "a little change";
                            dataGridView1[VALUE_COLUMN_NAME, list_index].Style.BackColor = Color.Yellow;
                            break;
                        default:
                            value_str = "CHANGE";
                            dataGridView1[VALUE_COLUMN_NAME, list_index].Style.BackColor = Color.Red;
                            row[CHANGED_COLUMN_NAME] = true;
                            ShowMessageOnBalloon("CHANGE", 30);
                            break;
                    }
                    row[VALUE_COLUMN_NAME] = value_str + " (" + distance.ToString() + ")";
                    row[DATE_COLUMN_NAME] = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                }
            }
        }

        private double CalculateHistogramDistance(string jpg_path1, string jpg_path2)
        {
            int i, sch = 0;
            float[] range_0 = { 0, 256 };
            float[][] ranges = { range_0 };
            double tmp, dist = 0;
            IplImage src_img1, src_img2;
            IplImage[] dst_img1 = new IplImage[4];
            IplImage[] dst_img2 = new IplImage[4];

            CvHistogram[] hist1 = new CvHistogram[4];
            CvHistogram hist2;

            src_img1 = IplImage.FromFile(jpg_path1, LoadMode.AnyDepth | LoadMode.AnyColor);

            // チャンネル数分の画像領域を確保
            sch = src_img1.NChannels;
            for (i = 0; i < sch; i++)
            {
                dst_img1[i] = Cv.CreateImage(Cv.Size(src_img1.Width, src_img1.Height), src_img1.Depth, 1);
            }
            // ヒストグラム構造体を確保
            int[] nHisSize = new int[1];
            nHisSize[0] = 256;
            hist1[0] = Cv.CreateHist(nHisSize, HistogramFormat.Array, ranges, true);

            // 入力画像がマルチチャンネルの場合，画像をチャンネル毎に分割
            if (sch == 1)
            {
                Cv.Copy(src_img1, dst_img1[0]);
            }
            else
            {
                Cv.Split(src_img1, dst_img1[0], dst_img1[1], dst_img1[2], dst_img1[3]);
            }

            for (i = 0; i < sch; i++)
            {
                Cv.CalcHist(dst_img1[i], hist1[i], false);
                Cv.NormalizeHist(hist1[i], 10000);
                if (i < 3)
                {
                    Cv.CopyHist(hist1[i], ref hist1[i + 1]);
                }
            }

            Cv.ReleaseImage(src_img1);

            src_img2 = IplImage.FromFile(jpg_path2, LoadMode.AnyDepth | LoadMode.AnyColor);

            // 入力画像のチャンネル数分の画像領域を確保
            for (i = 0; i < sch; i++)
            {
                dst_img2[i] = Cv.CreateImage(Cv.Size(src_img2.Width, src_img2.Height), src_img2.Depth, 1);
            }

            // ヒストグラム構造体を確保
            nHisSize[0] = 256;
            hist2 = Cv.CreateHist(nHisSize, HistogramFormat.Array, ranges, true);

            // 入力画像がマルチチャンネルの場合，画像をチャンネル毎に分割
            if (sch == 1)
            {
                Cv.Copy(src_img2, dst_img2[0]);
            }
            else
            {
                Cv.Split(src_img2, dst_img2[0], dst_img2[1], dst_img2[2], dst_img2[3]);
            }

            try
            {
                dist = 0.0;

                // ヒストグラムを計算，正規化して，距離を求める
                for (i = 0; i < sch; i++)
                {
                    Cv.CalcHist(dst_img2[i], hist2, false);
                    Cv.NormalizeHist(hist2, 10000);
                    tmp = Cv.CompareHist(hist1[i], hist2, HistogramComparison.Bhattacharyya);
                    dist += tmp * tmp;
                }
                dist = Math.Sqrt(dist);

                Cv.ReleaseHist(hist2);
                Cv.ReleaseImage(src_img2);
            }
            catch (OpenCVException ex)
            {
                Console.WriteLine("Error : " + ex.Message);
            }

            return dist;
        }

        private void RemoveCheckedRows()
        {
            dataTable.AsEnumerable().Where(row => ((bool)row[CHECK_COLUMN_NAME]) == true).ToList().ForEach(row => row.Delete());
        }

        private bool TryToDelete(string str)
        {
            try
            {
                System.IO.File.Delete(str);
            }
            catch (System.IO.IOException)
            {
                return false;
            }
            return true;
        }

        private void RemoveCheckedFiles()
        {
            var list = dataTable.AsEnumerable().Where(row => ((bool)row[CHECK_COLUMN_NAME]) == true).ToList();
            foreach (var row in list)
            {
                var filename_base = GetHash(row[URL_COLUMN_NAME].ToString());
                var file_path = FILE_PATH + filename_base;
                TryToDelete(file_path + OLD_FILE_SUFFIX);
                TryToDelete(file_path + NEW_FILE_SUFFIX);
            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (dataSet1.Tables[TABLE_NAME].Rows.Count != 0)
            {
                dataSet1.WriteXml(XML_PATH);
            }
        }

        private void dataGridView1_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == dataGridView1.Columns[URL_COLUMN_NAME].Index)
            {
                System.Diagnostics.Process.Start(dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString());
            }
        }

        private void addToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //Form2から送られてきたテキストを受け取る。
            ReceiveData data = Form2.ShowMiniForm();　//Form2を開く

            if (data.is_canceled == false)
            {
                dataSet1.Tables[TABLE_NAME].Rows.Add(false, data.url, "", "", false);
            }
        }

        private void deleteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
            RemoveCheckedFiles();
            RemoveCheckedRows();
        }

        private void checkToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (isChecking == false)
            {
                isChecking = true;
                progressBar1.Value = 1;

                StartCheck();
            }
        }

        private void initializeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            foreach (var file in System.IO.Directory.EnumerateFiles(FILE_PATH, "*.jpg"))
            {
                TryToDelete(file);
            }
            TryToDelete(XML_PATH);
            if (dataSet1.Tables[TABLE_NAME].Rows.Count != 0)
            {
                dataTable.AsEnumerable().ToList().ForEach(row => row.Delete());
            }
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == dataGridView1.Columns[CHECK_COLUMN_NAME].Index)
            {
                if ((bool)dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value == true)
                {
                    dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = false;
                }
                else
                {
                    dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = true;
                }
            }
        }
    }
}
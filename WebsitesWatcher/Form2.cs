using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WebsitesWatcher
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }
        private string[] argumentValues;
        public ReceiveData returnValue;

        public Form2(params string[] argumentValues)
        {
            this.argumentValues = argumentValues;

            InitializeComponent();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            this.returnValue.is_canceled = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.returnValue.url = textBox1.Text;
            this.returnValue.is_canceled = false;

            this.Close();
        }

        static public ReceiveData ShowMiniForm()
        {

            Form2 f = new Form2();
            f.ShowDialog();
            ReceiveData data = f.returnValue;
            f.Dispose();

            return data;
        }
    }
}

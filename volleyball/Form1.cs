using System;
using System.Windows.Forms;
using System.Drawing;

struct student
{
    public int no;//出席番号
    public string name;//名前

    public int sarve_s;//サーブ成功
    public int sarve_t;//サーブ合計数

    public int receive_s;//レシーブ成功
    public int receive_t;//レシーブ合計数

    public double par_s;//成功率(サーブ)
    public double par_r;//成功率(レシーブ)
}

namespace volleyball
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        student s1 = new student();
        student s2 = new student();
        student s3 = new student();
        student s4 = new student();
        student s5 = new student();
        student s6 = new student();
        student s7 = new student();

        /*-----------------見やすくするための補助線を描画---------------------*/
        private void Form1_Paint_1(object sender, PaintEventArgs e)
        {
            Graphics g = CreateGraphics();
            Pen pen = new Pen(Color.Black, 2);
            Pen lin = new Pen(Color.DarkGray, 2);
            g.DrawLine(lin, 444, 115, 444, Height);
            g.DrawLine(pen, 0, 115, Width, 115);
            g.DrawLine(pen, 0, 192, Width, 192);
            g.DrawLine(pen, 0, 255, Width, 255);
            g.DrawLine(pen, 0, 318, Width, 318);
            g.DrawLine(pen, 0, 381, Width, 381);
            g.DrawLine(pen, 0, 444, Width, 444);
            g.DrawLine(pen, 0, 507, Width, 507);
        }

        /*--------出席番号のボックスには数字のみ入力可能にする処理------------*/
        private void textBox1_KeyPress_1(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || '9' < e.KeyChar) && e.KeyChar != '\b')
            {
                e.Handled = true;
            }
        }

        private void textBox2_KeyPress_1(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || '9' < e.KeyChar) && e.KeyChar != '\b')
            {
                e.Handled = true;
            }
        }

        private void textBox3_KeyPress_1(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || '9' < e.KeyChar) && e.KeyChar != '\b')
            {
                e.Handled = true;
            }
        }

        private void textBox4_KeyPress_1(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || '9' < e.KeyChar) && e.KeyChar != '\b')
            {
                e.Handled = true;
            }
        }

        private void textBox5_KeyPress_1(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || '9' < e.KeyChar) && e.KeyChar != '\b')
            {
                e.Handled = true;
            }
        }

        private void textBox6_KeyPress_1(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || '9' < e.KeyChar) && e.KeyChar != '\b')
            {
                e.Handled = true;
            }
        }

        private void textBox7_KeyPress_1(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || '9' < e.KeyChar) && e.KeyChar != '\b')
            {
                e.Handled = true;
            }
        }

        /*---------------ボタンクリック時の処理-------------------------------*/
        /*-------サーブ成功--------------*/
        private void button1_Click(object sender, EventArgs e)//サーブ成功１
        {
            /** サーブ成功 **/
            s1.sarve_s++;
            s1.sarve_t++;
            s1.par_s = ((double)s1.sarve_s / s1.sarve_t) * 100;//サーブ率
            label9.Text = s1.sarve_s.ToString() + "回";
            label37.Text = s1.par_s.ToString("f0") + "%";
        }

        private void button5_Click(object sender, EventArgs e)//サーブ成功２
        {
            /** サーブ成功 **/
            s2.sarve_s++;
            s2.sarve_t++;
            s2.par_s = ((double)s2.sarve_s / s2.sarve_t) * 100;//サーブ率
            label13.Text = s2.sarve_s.ToString() + "回";
            label38.Text = s2.par_s.ToString("f0") + "%";
        }

        private void button9_Click(object sender, EventArgs e)//サーブ成功３
        {
            /** サーブ成功 **/
            s3.sarve_s++;
            s3.sarve_t++;
            s3.par_s = ((double)s3.sarve_s / s3.sarve_t) * 100;//サーブ率
            label17.Text = s3.sarve_s.ToString() + "回";
            label39.Text = s3.par_s.ToString("f0") + "%";
        }

        private void button13_Click(object sender, EventArgs e)//サーブ成功４
        {
            /** サーブ成功 **/
            s4.sarve_s++;
            s4.sarve_t++;
            s4.par_s = ((double)s4.sarve_s / s4.sarve_t) * 100;//サーブ率
            label21.Text = s4.sarve_s.ToString() + "回";
            label40.Text = s4.par_s.ToString("f0") + "%";
        }

        private void button17_Click(object sender, EventArgs e)//サーブ成功５
        {
            /** サーブ成功 **/
            s5.sarve_s++;
            s5.sarve_t++;
            s5.par_s = ((double)s5.sarve_s / s5.sarve_t) * 100;//サーブ率
            label25.Text = s5.sarve_s.ToString() + "回";
            label41.Text = s5.par_s.ToString("f0") + "%";
        }

        private void button21_Click(object sender, EventArgs e)//サーブ成功６
        {
            /** サーブ成功 **/
            s6.sarve_s++;
            s6.sarve_t++;
            s6.par_s = ((double)s6.sarve_s / s6.sarve_t) * 100;//サーブ率
            label29.Text = s6.sarve_s.ToString() + "回";
            label42.Text = s6.par_s.ToString("f0") + "%";
        }

        private void button25_Click(object sender, EventArgs e)//サーブ成功７
        {
            /** サーブ成功 **/
            s7.sarve_s++;
            s7.sarve_t++;
            s7.par_s = ((double)s7.sarve_s / s7.sarve_t) * 100;//サーブ率
            label33.Text = s7.sarve_s.ToString() + "回";
            label43.Text = s7.par_s.ToString("f0") + "%";
        }

        /*---------サーブ失敗-------------*/
        private void button2_Click(object sender, EventArgs e)//サーブ合計数１
        {
            /** サーブ合計数 **/
            s1.sarve_t++;
            s1.par_s = ((double)s1.sarve_s / s1.sarve_t) * 100;//サーブ率
            label37.Text = s1.par_s.ToString("f0") + "%";
            label10.Text = (s1.sarve_t - s1.sarve_s).ToString() + "回";
        }

        private void button6_Click(object sender, EventArgs e)//サーブ合計数２
        {
            /** サーブ合計数 **/
            s2.sarve_t++;
            s2.par_s = ((double)s2.sarve_s / s2.sarve_t) * 100;//サーブ率
            label38.Text = s2.par_s.ToString("f0") + "%";
            label14.Text = (s2.sarve_t - s2.sarve_s).ToString() + "回";
        }

        private void button10_Click(object sender, EventArgs e)//サーブ合計数３
        {
            /** サーブ合計数 **/
            s3.sarve_t++;
            s3.par_s = ((double)s3.sarve_s / s3.sarve_t) * 100;//サーブ率
            label39.Text = s3.par_s.ToString("f0") + "%";
            label18.Text = (s3.sarve_t - s3.sarve_s).ToString() + "回";
        }

        private void button14_Click(object sender, EventArgs e)//サーブ合計数４
        {
            /** サーブ合計数 **/
            s4.sarve_t++;
            s4.par_s = ((double)s4.sarve_s / s4.sarve_t) * 100;//サーブ率
            label40.Text = s4.par_s.ToString("f0") + "%";
            label22.Text = (s4.sarve_t - s4.sarve_s).ToString() + "回";
        }

        private void button18_Click(object sender, EventArgs e)//サーブ合計数５
        {
            /** サーブ合計数 **/
            s5.sarve_t++;
            s5.par_s = ((double)s5.sarve_s / s5.sarve_t) * 100;//サーブ率
            label41.Text = s5.par_s.ToString("f0") + "%";
            label26.Text = (s5.sarve_t - s5.sarve_s).ToString() + "回";
        }

        private void button22_Click(object sender, EventArgs e)//サーブ合計数６
        {
            /** サーブ合計数 **/
            s6.sarve_t++;
            s6.par_s = ((double)s6.sarve_s / s6.sarve_t) * 100;//サーブ率
            label42.Text = s6.par_s.ToString("f0") + "%";
            label30.Text = (s6.sarve_t - s6.sarve_s).ToString() + "回";
        }

        private void button26_Click(object sender, EventArgs e)//サーブ合計数７
        {
            /** サーブ合計数 **/
            s7.sarve_t++;
            s7.par_s = ((double)s7.sarve_s / s7.sarve_t) * 100;//サーブ率
            label43.Text = s7.par_s.ToString("f0") + "%";
            label34.Text = (s7.sarve_t - s7.sarve_s).ToString() + "回";
        }

        /*--------レシーブ成功------------*/
        private void button3_Click(object sender, EventArgs e)//レシーブ成功１
        {
            /** レシーブ成功 **/
            s1.receive_s++;//レシーブ成功＋１
            s1.receive_t++;
            s1.par_r = ((double)s1.receive_s / s1.receive_t) * 100;//サーブ率
            label11.Text = s1.receive_s.ToString() + "回";//レシーブ成功表示
            label44.Text = s1.par_r.ToString("f0") + "%";//成功率表示
        }

        private void button7_Click(object sender, EventArgs e)//レシーブ成功２
        {
            /** レシーブ成功 **/
            s2.receive_s++;//レシーブ成功＋１
            s2.receive_t++;
            s2.par_r = ((double)s2.receive_s / s2.receive_t) * 100;//サーブ率
            label15.Text = s2.receive_s.ToString() + "回";//レシーブ成功表示
            label45.Text = s2.par_r.ToString("f0") + "%";//成功率表示
        }

        private void button11_Click(object sender, EventArgs e)//レシーブ成功３
        {
            /** レシーブ成功 **/
            s3.receive_s++;//レシーブ成功＋１
            s3.receive_t++;
            s3.par_r = ((double)s3.receive_s / s3.receive_t) * 100;//サーブ率
            label19.Text = s3.receive_s.ToString() + "回";//レシーブ成功表示
            label46.Text = s3.par_r.ToString("f0") + "%";//成功率表示
        }

        private void button15_Click(object sender, EventArgs e)//レシーブ成功４
        {
            /** レシーブ成功 **/
            s4.receive_s++;//レシーブ成功＋１
            s4.receive_t++;
            s4.par_r = ((double)s4.receive_s / s4.receive_t) * 100;//サーブ率
            label23.Text = s4.receive_s.ToString() + "回";//レシーブ成功表示
            label47.Text = s4.par_r.ToString("f0") + "%";//成功率表示
        }

        private void button19_Click(object sender, EventArgs e)//レシーブ成功５
        {
            /** レシーブ成功 **/
            s5.receive_s++;//レシーブ成功＋１
            s5.receive_t++;
            s5.par_r = ((double)s5.receive_s / s5.receive_t) * 100;//サーブ率
            label27.Text = s5.receive_s.ToString() + "回";//レシーブ成功表示
            label48.Text = s5.par_r.ToString("f0") + "%";//成功率表示
        }

        private void button23_Click(object sender, EventArgs e)//レシーブ成功６
        {
            /** レシーブ成功 **/
            s6.receive_s++;//レシーブ成功＋１
            s6.receive_t++;
            s6.par_r = ((double)s6.receive_s / s6.receive_t) * 100;//サーブ率
            label31.Text = s6.receive_s.ToString() + "回";//レシーブ成功表示
            label49.Text = s6.par_r.ToString("f0") + "%";//成功率表示
        }

        private void button27_Click(object sender, EventArgs e)//レシーブ成功７
        {
            /** レシーブ成功 **/
            s7.receive_s++;//レシーブ成功＋１
            s7.receive_t++;
            s7.par_r = ((double)s7.receive_s / s7.receive_t) * 100;//サーブ率
            label35.Text = s7.receive_s.ToString() + "回";//レシーブ成功表示
            label50.Text = s7.par_r.ToString("f0") + "%";//成功率表示
        }

        /*--------レシーブ失敗------------*/
        private void button4_Click(object sender, EventArgs e)//レシーブ合計数１
        {
            /** レシーブ合計数 **/
            s1.receive_t++;//レシーブ合計数＋１
            s1.par_r = ((double)s1.receive_s / s1.receive_t) * 100;//レシーブ率
            label12.Text = (s1.receive_t - s1.receive_s).ToString() + "回";//レシーブ合計数表示
            label44.Text = s1.par_r.ToString("f0") + "%";
        }

        private void button8_Click(object sender, EventArgs e)//レシーブ合計数２
        {
            /** レシーブ合計数 **/
            s2.receive_t++;
            s2.par_r = ((double)s2.receive_s / s2.receive_t) * 100;//レシーブ率
            label16.Text = (s2.receive_t - s2.receive_s).ToString() + "回";
            label45.Text = s2.par_r.ToString("f0") + "%";
        }

        private void button12_Click(object sender, EventArgs e)//レシーブ合計数３
        {
            /** レシーブ合計数 **/
            s3.receive_t++;
            s3.par_r = ((double)s3.receive_s / s3.receive_t) * 100;//レシーブ率
            label20.Text = (s3.receive_t - s3.receive_s).ToString() + "回";
            label46.Text = s3.par_r.ToString("f0") + "%";
        }

        private void button16_Click(object sender, EventArgs e)//レシーブ合計数４
        {
            /** レシーブ合計数 **/
            s4.receive_t++;
            s4.par_r = ((double)s4.receive_s / s4.receive_t) * 100;//レシーブ率
            label24.Text = (s4.receive_t - s4.receive_s).ToString() + "回";
            label47.Text = s4.par_r.ToString("f0") + "%";
        }

        private void button20_Click(object sender, EventArgs e)//レシーブ合計数５
        {
            /** レシーブ合計数 **/
            s5.receive_t++;
            s5.par_r = ((double)s5.receive_s / s5.receive_t) * 100;//レシーブ率
            label28.Text = (s5.receive_t - s5.receive_s).ToString() + "回";
            label48.Text = s5.par_r.ToString("f0") + "%";
        }

        private void button24_Click(object sender, EventArgs e)//レシーブ合計数６
        {
            /** レシーブ合計数 **/
            s6.receive_t++;
            s6.par_r = ((double)s6.receive_s / s6.receive_t) * 100;//レシーブ率
            label32.Text = (s6.receive_t - s6.receive_s).ToString() + "回";
            label49.Text = s6.par_r.ToString("f0") + "%";
        }

        private void button28_Click(object sender, EventArgs e)//レシーブ合計数７
        {
            /** レシーブ合計数 **/
            s7.receive_t++;
            s7.par_r = ((double)s7.receive_s / s7.receive_t) * 100;//レシーブ率
            label36.Text = (s7.receive_t - s7.receive_s).ToString() + "回";
            label50.Text = s7.par_r.ToString("f0") + "%";
        }

        /*----------サーブ成功取り消し---------*/
        private void button30_Click(object sender, EventArgs e)
        {
            /** サーブ成功 **/
            s1.sarve_s--;
            s1.sarve_t--;
            s1.par_s = ((double)s1.sarve_s / s1.sarve_t) * 100;//サーブ率
            label9.Text = s1.sarve_s.ToString() + "回";
            label37.Text = s1.par_s.ToString("f0") + "%";
        }
        private void button34_Click(object sender, EventArgs e)
        {
            /** サーブ成功 **/
            s2.sarve_s--;
            s2.sarve_t--;
            s2.par_s = ((double)s2.sarve_s / s2.sarve_t) * 100;//サーブ率
            label13.Text = s2.sarve_s.ToString() + "回";
            label38.Text = s2.par_s.ToString("f0") + "%";
        }
        private void button38_Click(object sender, EventArgs e)
        {
            /** サーブ成功 **/
            s3.sarve_s--;
            s3.sarve_t--;
            s3.par_s = ((double)s3.sarve_s / s3.sarve_t) * 100;//サーブ率
            label17.Text = s3.sarve_s.ToString() + "回";
            label39.Text = s3.par_s.ToString("f0") + "%";
        }
        private void button42_Click(object sender, EventArgs e)
        {
            /** サーブ成功 **/
            s4.sarve_s--;
            s4.sarve_t--;
            s4.par_s = ((double)s4.sarve_s / s4.sarve_t) * 100;//サーブ率
            label21.Text = s4.sarve_s.ToString() + "回";
            label40.Text = s4.par_s.ToString("f0") + "%";
        }
        private void button46_Click(object sender, EventArgs e)
        {
            /** サーブ成功 **/
            s5.sarve_s--;
            s5.sarve_t--;
            s5.par_s = ((double)s5.sarve_s / s5.sarve_t) * 100;//サーブ率
            label25.Text = s5.sarve_s.ToString() + "回";
            label41.Text = s5.par_s.ToString("f0") + "%";
        }
        private void button50_Click(object sender, EventArgs e)
        {
            /** サーブ成功 **/
            s6.sarve_s--;
            s6.sarve_t--;
            s6.par_s = ((double)s6.sarve_s / s6.sarve_t) * 100;//サーブ率
            label29.Text = s6.sarve_s.ToString() + "回";
            label42.Text = s6.par_s.ToString("f0") + "%";
        }
        private void button54_Click(object sender, EventArgs e)
        {
            /** サーブ成功 **/
            s7.sarve_s--;
            s7.sarve_t--;
            s7.par_s = ((double)s7.sarve_s / s7.sarve_t) * 100;//サーブ率
            label33.Text = s7.sarve_s.ToString() + "回";
            label43.Text = s7.par_s.ToString("f0") + "%";
        }

        /*-----------サーブ失敗取り消し--------*/
        private void button31_Click(object sender, EventArgs e)
        {
            /** サーブ合計数 **/
            s1.sarve_t--;
            s1.par_s = ((double)s1.sarve_s / s1.sarve_t) * 100;//サーブ率
            label37.Text = s1.par_s.ToString("f0") + "%";
            label10.Text = (s1.sarve_t - s1.sarve_s).ToString() + "回";
        }
        private void button35_Click(object sender, EventArgs e)
        {
            /** サーブ合計数 **/
            s2.sarve_t--;
            s2.par_s = ((double)s2.sarve_s / s2.sarve_t) * 100;//サーブ率
            label38.Text = s2.par_s.ToString("f0") + "%";
            label14.Text = (s2.sarve_t - s2.sarve_s).ToString() + "回";
        }
        private void button39_Click(object sender, EventArgs e)
        {
            /** サーブ合計数 **/
            s3.sarve_t--;
            s3.par_s = ((double)s3.sarve_s / s3.sarve_t) * 100;//サーブ率
            label39.Text = s3.par_s.ToString("f0") + "%";
            label18.Text = (s3.sarve_t - s3.sarve_s).ToString() + "回";
        }
        private void button43_Click(object sender, EventArgs e)
        {
            /** サーブ合計数 **/
            s4.sarve_t--;
            s4.par_s = ((double)s4.sarve_s / s4.sarve_t) * 100;//サーブ率
            label40.Text = s4.par_s.ToString("f0") + "%";
            label22.Text = (s4.sarve_t - s4.sarve_s).ToString() + "回";
        }
        private void button47_Click(object sender, EventArgs e)
        {
            /** サーブ合計数 **/
            s5.sarve_t--;
            s5.par_s = ((double)s5.sarve_s / s5.sarve_t) * 100;//サーブ率
            label41.Text = s5.par_s.ToString("f0") + "%";
            label26.Text = (s5.sarve_t - s5.sarve_s).ToString() + "回";
        }
        private void button51_Click(object sender, EventArgs e)
        {
            /** サーブ合計数 **/
            s6.sarve_t--;
            s6.par_s = ((double)s6.sarve_s / s6.sarve_t) * 100;//サーブ率
            label42.Text = s6.par_s.ToString("f0") + "%";
            label30.Text = (s6.sarve_t - s6.sarve_s).ToString() + "回";
        }
        private void button55_Click(object sender, EventArgs e)
        {
            /** サーブ合計数 **/
            s7.sarve_t--;
            s7.par_s = ((double)s7.sarve_s / s7.sarve_t) * 100;//サーブ率
            label43.Text = s7.par_s.ToString("f0") + "%";
            label34.Text = (s7.sarve_t - s7.sarve_s).ToString() + "回";
        }

        /*----------レシーブ成功取り消し-------*/
        private void button32_Click(object sender, EventArgs e)
        {
            /** レシーブ成功 **/
            s1.receive_s--;//レシーブ成功＋１
            s1.receive_t--;
            s1.par_r = ((double)s1.receive_s / s1.receive_t) * 100;//サーブ率
            label11.Text = s1.receive_s.ToString() + "回";//レシーブ成功表示
            label44.Text = s1.par_r.ToString("f0") + "%";//成功率表示
        }
        private void button36_Click(object sender, EventArgs e)
        {
            /** レシーブ成功 **/
            s2.receive_s--;//レシーブ成功＋１
            s2.receive_t--;
            s2.par_r = ((double)s2.receive_s / s2.receive_t) * 100;//サーブ率
            label15.Text = s2.receive_s.ToString() + "回";//レシーブ成功表示
            label45.Text = s2.par_r.ToString("f0") + "%";//成功率表示
        }
        private void button40_Click(object sender, EventArgs e)
        {
            /** レシーブ成功 **/
            s3.receive_s--;//レシーブ成功＋１
            s3.receive_t--;
            s3.par_r = ((double)s3.receive_s / s3.receive_t) * 100;//サーブ率
            label19.Text = s3.receive_s.ToString() + "回";//レシーブ成功表示
            label46.Text = s3.par_r.ToString("f0") + "%";//成功率表示
        }
        private void button44_Click(object sender, EventArgs e)
        {
            /** レシーブ成功 **/
            s4.receive_s--;//レシーブ成功＋１
            s4.receive_t--;
            s4.par_r = ((double)s4.receive_s / s4.receive_t) * 100;//サーブ率
            label23.Text = s4.receive_s.ToString() + "回";//レシーブ成功表示
            label47.Text = s4.par_r.ToString("f0") + "%";//成功率表示
        }
        private void button48_Click(object sender, EventArgs e)
        {
            /** レシーブ成功 **/
            s5.receive_s--;//レシーブ成功＋１
            s5.receive_t--;
            s5.par_r = ((double)s5.receive_s / s5.receive_t) * 100;//サーブ率
            label27.Text = s5.receive_s.ToString() + "回";//レシーブ成功表示
            label48.Text = s5.par_r.ToString("f0") + "%";//成功率表示
        }
        private void button52_Click(object sender, EventArgs e)
        {
            /** レシーブ成功 **/
            s6.receive_s--;//レシーブ成功＋１
            s6.receive_t--;
            s6.par_r = ((double)s6.receive_s / s6.receive_t) * 100;//サーブ率
            label31.Text = s6.receive_s.ToString() + "回";//レシーブ成功表示
            label49.Text = s6.par_r.ToString("f0") + "%";//成功率表示
        }
        private void button56_Click(object sender, EventArgs e)
        {
            /** レシーブ成功 **/
            s7.receive_s--;//レシーブ成功＋１
            s7.receive_t--;
            s7.par_r = ((double)s7.receive_s / s7.receive_t) * 100;//サーブ率
            label35.Text = s7.receive_s.ToString() + "回";//レシーブ成功表示
            label50.Text = s7.par_r.ToString("f0") + "%";//成功率表示
        }

        /*----------レシーブ失敗取り消し------*/
        private void button33_Click(object sender, EventArgs e)
        {
            /** レシーブ合計数 **/
            s1.receive_t--;//レシーブ合計数＋１
            s1.par_r = ((double)s1.receive_s / s1.receive_t) * 100;//レシーブ率
            label12.Text = (s1.receive_t - s1.receive_s).ToString() + "回";//レシーブ合計数表示
            label44.Text = s1.par_r.ToString("f0") + "%";
        }
        private void button37_Click(object sender, EventArgs e)
        {
            /** レシーブ合計数 **/
            s2.receive_t--;
            s2.par_r = ((double)s2.receive_s / s2.receive_t) * 100;//レシーブ率
            label16.Text = (s2.receive_t - s2.receive_s).ToString() + "回";
            label45.Text = s2.par_r.ToString("f0") + "%";
        }
        private void button41_Click(object sender, EventArgs e)
        {
            /** レシーブ合計数 **/
            s3.receive_t--;
            s3.par_r = ((double)s3.receive_s / s3.receive_t) * 100;//レシーブ率
            label20.Text = (s3.receive_t - s3.receive_s).ToString() + "回";
            label46.Text = s3.par_r.ToString("f0") + "%";
        }
        private void button45_Click(object sender, EventArgs e)
        {
            /** レシーブ合計数 **/
            s4.receive_t--;
            s4.par_r = ((double)s4.receive_s / s4.receive_t) * 100;//レシーブ率
            label24.Text = (s4.receive_t - s4.receive_s).ToString() + "回";
            label47.Text = s4.par_r.ToString("f0") + "%";
        }
        private void button49_Click(object sender, EventArgs e)
        {
            /** レシーブ合計数 **/
            s5.receive_t--;
            s5.par_r = ((double)s5.receive_s / s5.receive_t) * 100;//レシーブ率
            label28.Text = (s5.receive_t - s5.receive_s).ToString() + "回";
            label48.Text = s5.par_r.ToString("f0") + "%";
        }
        private void button53_Click(object sender, EventArgs e)
        {
            /** レシーブ合計数 **/
            s6.receive_t--;
            s6.par_r = ((double)s6.receive_s / s6.receive_t) * 100;//レシーブ率
            label32.Text = (s6.receive_t - s6.receive_s).ToString() + "回";
            label49.Text = s6.par_r.ToString("f0") + "%";
        }
        private void button57_Click(object sender, EventArgs e)
        {
            /** レシーブ合計数 **/
            s7.receive_t--;
            s7.par_r = ((double)s7.receive_s / s7.receive_t) * 100;//レシーブ率
            label36.Text = (s7.receive_t - s7.receive_s).ToString() + "回";
            label50.Text = s7.par_r.ToString("f0") + "%";
        }

        /*-------------確定ボタン，エクセルに出力-----------------*/
        private void button29_Click(object sender, EventArgs e)
        {
            /* データの完成 */
            string grad = comboBox1.Text;
            string gakka = comboBox2.Text;
            string team = comboBox3.Text;
            if (textBox1.Text != "")
            {
                s1.no = Convert.ToInt32(textBox1.Text);
            }
            if (textBox2.Text != "")
            {
                s2.no = Convert.ToInt32(textBox2.Text);
            }
            if (textBox3.Text != "")
            {
                s3.no = Convert.ToInt32(textBox3.Text);
            }
            if (textBox4.Text != "")
            {
                s4.no = Convert.ToInt32(textBox4.Text);
            }
            if (textBox5.Text != "")
            {
                s5.no = Convert.ToInt32(textBox5.Text);
            }
            if (textBox6.Text != "")
            {
                s6.no = Convert.ToInt32(textBox6.Text);
            }
            if (textBox7.Text != "")
            {
                s7.no = Convert.ToInt32(textBox7.Text);
            }
            s1.name = textBox8.Text;
            s2.name = textBox9.Text;
            s3.name = textBox10.Text;
            s4.name = textBox11.Text;
            s5.name = textBox12.Text;
            s6.name = textBox13.Text;
            s7.name = textBox14.Text;

            /* 保存の準備 */
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Title = "保存先の指定";// ダイアログのタイトルを設定する
            saveFileDialog1.InitialDirectory = @"C:\";// 初期表示するディレクトリを設定する
            saveFileDialog1.FileName = "学年学科チーム名を入力";// 初期表示するファイル名を設定する
            saveFileDialog1.Filter = "Excel ファイル(*.xlsx *.xls)|*.xlsx;*.xls|すべてのファイル(*.*)|*.*";// ファイルのフィルタを設定する
            saveFileDialog1.RestoreDirectory = true;// ダイアログボックスを閉じる前に現在のディレクトリを復元する (初期値 false)
            saveFileDialog1.ShowHelp = true;// [ヘルプ] ボタンを表示する (初期値 false)
            saveFileDialog1.CreatePrompt = true;// 存在しないファイルを指定した場合は、新しく作成するかどうかの問い合わせを表示する (初期値 false)
            saveFileDialog1.AddExtension = true;// 拡張子を指定しない場合は自動的に拡張子を付加する (初期値 true)
            saveFileDialog1.ValidateNames = true;// 有効な Win32 ファイル名だけを受け入れるようにする (初期値 true)
            saveFileDialog1.Dispose();// 不要になった時点で破棄する (正しくは オブジェクトの破棄を保証する を参照)
            if (saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                label6.Text = saveFileDialog1.FileName;// 保存先を表示
            }
            string ExcelBookFileName = label6.Text;// ファイル名決定

            var xls = new CKExcelLB.ExcelLB(); // ExcelLBのインスタンス作成 (Excelオブジェクト作成)
            xls.Visible = true;
            var books = xls.Workbooks;
            var book1 = xls.AddBook(books);
            var sheets = xls.GetSheets(book1);
            var sheet1 = xls.GetSheet(sheets, 1);

            label7.Text = "完了までしばらくお待ちください";

            /* エクセルに書き込む内容 */
            //項目名
            xls.SetRange(sheet1, 2, 1, "学年");//A
            xls.SetRange(sheet1, 2, 2, "学科");//B
            xls.SetRange(sheet1, 2, 3, "チーム");//C
            xls.SetRange(sheet1, 2, 4, "番号");//D
            xls.SetRange(sheet1, 2, 5, "氏名");//E
            xls.SetRange(sheet1, 2, 6, "サーブ成功");//F
            xls.SetRange(sheet1, 2, 7, "サーブ失敗");//G
            xls.SetRange(sheet1, 2, 8, "サーブ合計");//H
            xls.SetRange(sheet1, 2, 9, "サーブ成功率");//I
            xls.SetRange(sheet1, 2, 10, "レシーブ成功");//J
            xls.SetRange(sheet1, 2, 11, "レシーブ失敗");//K
            xls.SetRange(sheet1, 2, 12, "レシーブ合計");//L
            xls.SetRange(sheet1, 2, 13, "レシーブ成功率");//M
            //学年:A
            xls.SetRange(sheet1, "A3:A9", grad);
            //学科:B
            xls.SetRange(sheet1, "B3:B9", gakka);
            //チーム:C
            xls.SetRange(sheet1, "C3:C9", team);
            //出席番号:D
            xls.SetRange(sheet1, 3, 4, s1.no);
            xls.SetRange(sheet1, 4, 4, s2.no);
            xls.SetRange(sheet1, 5, 4, s3.no);
            xls.SetRange(sheet1, 6, 4, s4.no);
            xls.SetRange(sheet1, 7, 4, s5.no);
            xls.SetRange(sheet1, 8, 4, s6.no);
            xls.SetRange(sheet1, 9, 4, s7.no);
            //氏名:E
            xls.SetRange(sheet1, 3, 5, s1.name);
            xls.SetRange(sheet1, 4, 5, s2.name);
            xls.SetRange(sheet1, 5, 5, s3.name);
            xls.SetRange(sheet1, 6, 5, s4.name);
            xls.SetRange(sheet1, 7, 5, s5.name);
            xls.SetRange(sheet1, 8, 5, s6.name);
            xls.SetRange(sheet1, 9, 5, s7.name);
            //サーブ成功:F
            xls.SetRange(sheet1, 3, 6, s1.sarve_s);
            xls.SetRange(sheet1, 4, 6, s2.sarve_s);
            xls.SetRange(sheet1, 5, 6, s3.sarve_s);
            xls.SetRange(sheet1, 6, 6, s4.sarve_s);
            xls.SetRange(sheet1, 7, 6, s5.sarve_s);
            xls.SetRange(sheet1, 8, 6, s6.sarve_s);
            xls.SetRange(sheet1, 9, 6, s7.sarve_s);
            //サーブ失敗:G
            xls.SetRange(sheet1, 3, 7, s1.sarve_t - s1.sarve_s);
            xls.SetRange(sheet1, 4, 7, s2.sarve_t - s2.sarve_s);
            xls.SetRange(sheet1, 5, 7, s3.sarve_t - s3.sarve_s);
            xls.SetRange(sheet1, 6, 7, s4.sarve_t - s4.sarve_s);
            xls.SetRange(sheet1, 7, 7, s5.sarve_t - s5.sarve_s);
            xls.SetRange(sheet1, 8, 7, s6.sarve_t - s6.sarve_s);
            xls.SetRange(sheet1, 9, 7, s7.sarve_t - s7.sarve_s);
            //サーブ合計:H
            xls.SetRange(sheet1, 3, 8, s1.sarve_t);
            xls.SetRange(sheet1, 4, 8, s2.sarve_t);
            xls.SetRange(sheet1, 5, 8, s3.sarve_t);
            xls.SetRange(sheet1, 6, 8, s4.sarve_t);
            xls.SetRange(sheet1, 7, 8, s5.sarve_t);
            xls.SetRange(sheet1, 8, 8, s6.sarve_t);
            xls.SetRange(sheet1, 9, 8, s7.sarve_t);
            //サーブ成功率:I
            xls.SetRange(sheet1, 3, 9, "=F3/H3");
            xls.SetRange(sheet1, 4, 9, "=F4/H4");
            xls.SetRange(sheet1, 5, 9, "=F5/H5");
            xls.SetRange(sheet1, 6, 9, "=F6/H6");
            xls.SetRange(sheet1, 7, 9, "=F7/H7");
            xls.SetRange(sheet1, 8, 9, "=F8/H8");
            xls.SetRange(sheet1, 9, 9, "=F9/H9");
            //レシーブ成功:J
            xls.SetRange(sheet1, 3, 10, s1.receive_s);
            xls.SetRange(sheet1, 4, 10, s2.receive_s);
            xls.SetRange(sheet1, 5, 10, s3.receive_s);
            xls.SetRange(sheet1, 6, 10, s4.receive_s);
            xls.SetRange(sheet1, 7, 10, s5.receive_s);
            xls.SetRange(sheet1, 8, 10, s6.receive_s);
            xls.SetRange(sheet1, 9, 10, s7.receive_s);
            //レシーブ失敗:K
            xls.SetRange(sheet1, 3, 11, s1.receive_t - s1.receive_s);
            xls.SetRange(sheet1, 4, 11, s2.receive_t - s2.receive_s);
            xls.SetRange(sheet1, 5, 11, s3.receive_t - s3.receive_s);
            xls.SetRange(sheet1, 6, 11, s4.receive_t - s4.receive_s);
            xls.SetRange(sheet1, 7, 11, s5.receive_t - s5.receive_s);
            xls.SetRange(sheet1, 8, 11, s6.receive_t - s6.receive_s);
            xls.SetRange(sheet1, 9, 11, s7.receive_t - s7.receive_s);
            //レシーブ合計;L
            xls.SetRange(sheet1, 3, 12, s1.receive_t);
            xls.SetRange(sheet1, 4, 12, s2.receive_t);
            xls.SetRange(sheet1, 5, 12, s3.receive_t);
            xls.SetRange(sheet1, 6, 12, s4.receive_t);
            xls.SetRange(sheet1, 7, 12, s5.receive_t);
            xls.SetRange(sheet1, 8, 12, s6.receive_t);
            xls.SetRange(sheet1, 9, 12, s7.receive_t);
            //レシーブ成功率:M
            xls.SetRange(sheet1, 3, 13, "=J3/L3");
            xls.SetRange(sheet1, 4, 13, "=J4/L4");
            xls.SetRange(sheet1, 5, 13, "=J5/L5");
            xls.SetRange(sheet1, 6, 13, "=J6/L6");
            xls.SetRange(sheet1, 7, 13, "=J7/L7");
            xls.SetRange(sheet1, 8, 13, "=J8/L8");
            xls.SetRange(sheet1, 9, 13, "=J9/L9");

            // A～M列の幅を自動調整
            xls.AutoFitColumnWidth(sheet1, "A:M");

            // bookの保存
            xls.DisplayAlerts = false;
            /* エクセルの保存と終了 */
            label7.Text = "";
           
            xls.SaveAs(book1, ExcelBookFileName);
            // COMオブジェクトの解放
            xls.ReleaseObject(sheet1);
            xls.ReleaseObject(sheets);
            xls.ReleaseObject(book1);
            xls.ReleaseObject(books);

            xls.DisplayAlerts = true;
            xls.Quit();
            xls.Dispose();
            MessageBox.Show("保存が完了しました．", "完了");
            if (ExcelBookFileName == "")
            {
                MessageBox.Show("保存ができていない可能性があります．", "失敗", MessageBoxButtons.OK, MessageBoxIcon.Hand);
            }
        }

    }
}

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Marksheet
{
    public partial class Common_Details : Form
    {
        string testno,sem, branch, testperiod;
        DataSheet mark;
        public Common_Details()
        {
            InitializeComponent();
            testno = "I";
            testperiod = dateTimePicker1.Value.ToShortDateString();
            sem = "I";
            branch = "Computer Science ";
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if(sem.CompareTo("I")==0||sem.CompareTo("II")==0)
                mark = new DataSheet(testno,"I", sem, branch, testperiod);
            if (sem.CompareTo("III") == 0 || sem.CompareTo("IV") == 0)
                mark = new DataSheet(testno, "II", sem, branch, testperiod);
            if (sem.CompareTo("V") == 0 || sem.CompareTo("VI") == 0)
                mark = new DataSheet(testno, "III", sem, branch, testperiod);
            if (sem.CompareTo("VII") == 0 || sem.CompareTo("VIII") == 0)
                mark = new DataSheet(testno, "IV", sem, branch, testperiod);
            this.Hide();
            mark.ShowDialog();
            this.Close();

        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if(radioButton1.Checked==true)
                testno = "I";
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton3.Checked == true)
               testno = "II";
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if(radioButton2.Checked==true)
              testno = "II";
        }

        private void radioButton11_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton11.Checked == true)
                sem = "I";
        }

        private void radioButton12_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton12.Checked == true)
                sem = "II";
        }

        private void radioButton13_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton13.Checked == true)
                sem = "III";
        }

        private void radioButton14_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton14.Checked == true)
                sem = "IV";
        }

        private void radioButton15_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton15.Checked == true)
                sem = "V";
        }

        private void radioButton17_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton17.Checked == true)
                sem = "VI";
        }

        private void radioButton16_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton16.Checked == true)
                sem = "VII";
        }

        private void radioButton18_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton18.Checked == true)
                sem = "VIII";
        }

        private void radioButton6_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton6.Checked == true)
                branch="AutoMobile";
        }

        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton4.Checked == true)
                branch = "Computer Science";
        }

        private void radioButton5_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton5.Checked == true)
                branch = "Information Technology";
        }

        private void radioButton7_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton7.Checked == true)
                branch = "Civil Engineering";
        }

        private void radioButton8_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton8.Checked == true)
                branch = "Mechanical Engineering";
        }

        private void radioButton10_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton10.Checked == true)
                branch = "Electronics&Electrical Engineering";
        }

        private void radioButton9_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton9.Checked == true)
                branch = "Electronics&Communication Engineering";
        }
    }
}

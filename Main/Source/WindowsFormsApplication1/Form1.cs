using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using EInvoice_Chevron;
namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            EInvoice einvoice = new EInvoice();
            int DelTick;
            DelTick = 2147390252;
                //Chesapeake: 2147352613;
            einvoice.Main(@"D:\EDIInv\" + Convert.ToString(txtDelTicket.Text), Convert.ToInt32(txtDelTicket.Text));

        }
    }
}

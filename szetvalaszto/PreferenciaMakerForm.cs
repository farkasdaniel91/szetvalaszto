using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace szetvalaszto
{
    public partial class PreferenciaMakerForm : Form
    {
        public string Picker;
        public List<Par> Parok;
        public List<Par> ValaszthatoParok;
        public int PreferenciaPontok;
        public List<Preferencia> Preferenciak;
        public PreferenciaMakerForm(string picker, List<Par> parok)
        {
            InitializeComponent();
            this.Picker = picker;
            this.Parok = parok;
            this.ValaszthatoParok = parok;
            this.Preferenciak = new List<Preferencia>();
            var asd = ConfigurationManager.AppSettings["PreferenciaPontok"];
            this.PreferenciaPontok = Convert.ToInt32(asd);
        }

        private void Form2_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            this.comboBox1.DataSource = this.ValaszthatoParok.Select(x => x.par).ToArray();
            this.label1.Text = this.PreferenciaPontok.ToString();

            RefreshPrefPontz();
            this.Text = this.Picker;
        }

        private void RefreshPrefPontz()
        {
            this.comboBox2.Items.Clear();
            for (int i = 1; i < this.PreferenciaPontok + 1; i++)
            {
                this.comboBox2.Items.Add(i);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (this.comboBox1.SelectedItem == null || this.comboBox2.SelectedItem == null)
            {
                return;
            }

            string valasztott = this.comboBox1.SelectedItem.ToString();
            int prefpont = Convert.ToInt32(this.comboBox2.SelectedItem.ToString());
            string key = valasztott + " " + prefpont;

            if (this.PreferenciaPontok - prefpont < 0)
            {
                return;
            }

            this.listBox1.Items.Add(key);

            Preferencia pref = new Preferencia(this.Picker, valasztott, prefpont, key);
            this.Preferenciak.Add(pref);
            
            this.ValaszthatoParok.Remove(this.ValaszthatoParok.Where(x => x.par == comboBox1.SelectedItem.ToString()).First());
            this.comboBox1.DataSource = null;
            this.comboBox1.DataSource = this.ValaszthatoParok.Select(x => x.par).ToArray();

            this.PreferenciaPontok -= prefpont;
            label1.Text = this.PreferenciaPontok.ToString();
            RefreshPrefPontz();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (this.listBox1.SelectedItem == null)
            {
                return;
            }

            string key = this.listBox1.SelectedItem.ToString();
            Preferencia selectedPref = this.Preferenciak.Where(x => x.key == key).First();

            this.ValaszthatoParok.Add(new Par(selectedPref.valasztott));
            this.comboBox1.DataSource = null;
            this.comboBox1.DataSource = this.ValaszthatoParok.Select(x => x.par).ToArray();

            this.PreferenciaPontok += selectedPref.prefpont;

            this.listBox1.Items.Remove(key);
            this.Preferenciak.Remove(selectedPref);

            label1.Text = this.PreferenciaPontok.ToString();
            RefreshPrefPontz();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Biztos hogy jól döntesz? ne legyél noob..", "noob vagy", MessageBoxButtons.YesNo) != System.Windows.Forms.DialogResult.Yes)
	        {
                return;
	        }

            SzetvalasztoHelper.ExportPreferenciak(this.Preferenciak);

            MessageBox.Show("Te vagy a kedvenc instruktorom!", "tegeci", MessageBoxButtons.OK, MessageBoxIcon.Information);
            this.Close();
        }
    }
}

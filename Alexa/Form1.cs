using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using HtmlAgilityPack;

namespace Alexa
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            string Url = "http://www.alexa.com/topsites/countries";
            HtmlWeb web = new HtmlWeb();
            HtmlAgilityPack.HtmlDocument doc = web.Load(Url);
            List<ComboboxItem> item = new List<ComboboxItem>();
            item.Add(new ComboboxItem { Text = "Global", Value = "/topsites/global" });
            foreach (HtmlNode link in doc.DocumentNode.SelectNodes("//div[@class='categories top']//ul//li//a"))
            {
                item.Add(new ComboboxItem { Text = link.InnerText, Value = link.Attributes["href"].Value });
            }
            comboBox1.DataSource = new BindingSource(item, null);
        }
        List<TopSites> someStringList = new List<TopSites>();
        void AddToList(params string[] list)
        {
            for (int i = 0; i < list.Length; i++)
            {
                someStringList.Add(new TopSites { c = list[i] });
            }
        }
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            button1.Enabled = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SaveFileDialog sv = new SaveFileDialog();
            sv.Filter= "CSV (*.csv)|*.csv";
            sv.FileName = (comboBox1.SelectedValue as ComboboxItem).Text.ToString();
            var a = sv.ShowDialog();
            if (a == DialogResult.OK)
            {
                string Url = "http://www.alexa.com";
                HtmlWeb web = new HtmlWeb();
                List<TopSites> t = new List<TopSites>();
                for (int j = 0; j < 20; j++)
                {
                    if ((comboBox1.SelectedValue as ComboboxItem).Text != "Global")
                    {
                        HtmlAgilityPack.HtmlDocument doc = web.Load(Url + (comboBox1.SelectedValue as ComboboxItem).Value.ToString().Replace("countries", "countries;" + j.ToString()));
                        HtmlNode[] count = doc.DocumentNode.SelectNodes("//div[@class='listings']//ul//li//div[@class='count']").ToArray();
                        HtmlNode[] href = doc.DocumentNode.SelectNodes("//div[@class='listings']//ul//li//p[@class='desc-paragraph']//a").ToArray();
                        HtmlNode[] desc = doc.DocumentNode.SelectNodes("//div[@class='listings']//ul//li//div[@class='description']").ToArray();
                        for (int i = 0; i < 25; i++)
                        {
                            t.Add(new TopSites { c = ((HtmlNode)count[i]).InnerText, h = ((HtmlNode)href[i]).InnerText, d = "\""+((HtmlNode)desc[i]).InnerText.Replace("\n"," ").Replace("\r"," ").Replace("&nbsp;", " ").Replace("More", " ").Replace("&hellip;", " ") + "\"" });
                        }
                    }
                    else
                    {
                        HtmlAgilityPack.HtmlDocument doc = web.Load(Url + (comboBox1.SelectedValue as ComboboxItem).Value.ToString().Replace("global", "global;" + j.ToString()));
                        HtmlNode[] count = doc.DocumentNode.SelectNodes("//div[@class='listings']//ul//li//div[@class='count']").ToArray();
                        HtmlNode[] href = doc.DocumentNode.SelectNodes("//div[@class='listings']//ul//li//p[@class='desc-paragraph']//a").ToArray();
                        HtmlNode[] desc = doc.DocumentNode.SelectNodes("//div[@class='listings']//ul//li//div[@class='description']").ToArray();
                        for (int i = 0; i < 25; i++)
                        {
                            t.Add(new TopSites { c = ((HtmlNode)count[i]).InnerText, h = ((HtmlNode)href[i]).InnerText, d = "\"" + ((HtmlNode)desc[i]).InnerText.Replace("\n", " ").Replace("\r", " ").Replace("&nbsp;"," ").Replace("More"," ").Replace("&hellip;", " ") + "\"" });
                        }
                    }   
                }
                using (var sw = new StreamWriter(File.Open(sv.FileName, FileMode.CreateNew), Encoding.UTF8))
                {
                    for (int index = 0; index < t.Count; index++)
                    {
                        sw.WriteLine(t[index].c + "," + t[index].h + "," + t[index].d);
                    }
                }
                MessageBox.Show("فایل مورد نظر با موفقیت ایجاد شد");
                button1.Enabled = false;
            }
        }
    }
    public class ComboboxItem
    {
        public string Text { get; set; }
        public object Value { get; set; }

        public override string ToString()
        {
            return Text;
        }
    }
    public class TopSites
    {
        public string c { get; set; }
        public string h { get; set; }
        public string d { get; set; }
    }

}

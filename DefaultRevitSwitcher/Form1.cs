using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Text.RegularExpressions;

namespace DefaultRevitSwitcher
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
     

            // Search Installed Revit Versions
            try
            {

                ComboboxItem item = new ComboboxItem();
                string LocalDrive = Path.GetPathRoot(Environment.SystemDirectory);
                string command =" /dde";
                string CurrentRevit = @"Program Files\Autodesk\Revit 20";
                int maxYear = Convert.ToInt32(DateTime.Now.AddYears(1).ToString("yy"));

                for (int i = 0; i < maxYear; i++)
                {
                    String Revit = LocalDrive + CurrentRevit + i.ToString() + @"\Revit.exe";

                    if (File.Exists(Revit))
                    {

                        comboBox1.Items.Add(new ComboboxItem { Name = "Revit 20" + i.ToString(), Value = '"' + Revit + '"' + command});
                        comboBox1.DisplayMember = "Name";
                        comboBox1.ValueMember = "Value";
                        comboBox1.SelectedIndex = 0;
                    }
   
                }


                String registryKey = @"Software\Classes\Revit.Project\shell\open\command";
                using (RegistryKey key = Registry.LocalMachine.OpenSubKey(registryKey))
                {
                    if (key != null)
                    {
                       // MessageBox.Show(key.GetValue("").ToString());
                        Regex regex = new Regex(@"\bRevit+ (?:\d*\.)?\d+");
                        Match match = regex.Match(key.GetValue("").ToString());
                        if (match.Success)
                        {
                            label2.Text = match.Value;
                        }  
                        
                    }
                }
            }
            catch (Exception ex)  
            {
                MessageBox.Show(ex.Message);

            }

            
        }
        public class ComboboxItem
        {
            public string Name { get; set; }
            public string Value { get; set; }

            public override string ToString()
            {
                return Name;
            }
        }
        private ComboboxItem selectedCar;
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox cmb = (ComboBox)sender;
            int selectedIndex = cmb.SelectedIndex;
           
            string selectedValue = (string)cmb.SelectedValue;

             selectedCar = (ComboboxItem)cmb.SelectedItem;


            //MessageBox.Show(selectedCar.Value); 
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {


                String registryKey = @"Software\Classes\Revit.Project\shell\open\command";
                using (RegistryKey key = Registry.LocalMachine.OpenSubKey(registryKey, true))
                {
                    if (key != null)
                    {
                        key.SetValue("", selectedCar.Value);
                        this.Close();
                    }
                
            }
        }
    }
}

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
using IronXL;
using Stimulsoft.Report;

namespace KartKanoon
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            Stimulsoft.Base.StiLicense.LoadFromString("6vJhGtLLLz2GNviWmUTrhSqnOItdDwjBylQzQcAOiHkcgIvwL0jnpsDqRpWg5FI5kt2G7A0tYIcUygBh1sPs7uPvgjp0GgDowCB/F6myz180QOXN+hAWpeqWhPokj7sFSjITHge+0Hvjw4vKQPmlfDn/OWCMfhCPY4cZMTeUW6cW2VSK+480C7TeIrX/O/tLgGrgklP1P/7MdEkP/gQjQIwyRizsmj17wLkWfRzMal1duePiYgMsYr8GE9AdT2Mz6RPH+SCwPKHdjCq5PI/k4SrswBNYyd60U3YHOve2dNPfteBnaTnzwpyfuKQSyJrPuccoqDVxIUWSF8GCXtQa2nf7qqvv7A9L4L2LSU3JS31y3NU4ykT1r2gg8lkLmXQlauRyq3SR3zhTCvr1gsuM8a/85YPA2KCT4T2X14/Sj6Z3uo9x8lFQPOsW3fk1us4HDqN54uz7DOynURHiLJ5Twy7m2SzZhgg7QKO07CZgff70N6ID1D/h2G8pjVhsUO5qkWEkdr2kj8ygbUq5OZcMYTuQXkt1+sVOet7/cmQGdjsxperXlhn/96fbzPPn/q4Q");
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            StiReport report = new StiReport();
            report.RegData("Kanooni", readExecl(@".\Content\Excel\kart.xlsx"));
            report.Load(@".\Content\StimulsoftDesigne\KDZKart.mrt");
            report.Show();
        }

        private KDZMember[] readExecl(string execlFile)
        {
            var workBook = WorkBook.Load(execlFile);
            var sheet = workBook.WorkSheets.FirstOrDefault();
            var members = new List<KDZMember>();
            for (int i = 2; i < 27; i++)
            {
                var name = sheet[$"A{i}"].StringValue;
                var fatherName = sheet[$"E{i}"].StringValue;
                var number = sheet[$"H{i}"].StringValue;
                var nationalCode = sheet[$"B{i}"].StringValue;
                var reshte = sheet[$"F{i}"].StringValue;
                var tavalod = sheet[$"C{i}"].StringValue;
                var university = sheet[$"G{i}"].StringValue;
                var ax = $"C:\\Users\\behro\\Pictures\\AxKDZ\\{sheet[$"K{i}"]}";
                members.Add(new KDZMember(name,fatherName,number,nationalCode,reshte,tavalod,university,ax));
            }

            return members.ToArray();
        }
    }

    internal class KDZMember
    {
        public string FullName { get; set; }
        public string FatherName { get; set; }
        public string Number { get; set; }
        public string NationalCode { get; set; }
        public string Reshte { get; set; }
        public string Tavalod { get; set; }
        public string University { get; set; }
        public Image ImageFile { get; set; }

        public KDZMember(string fullName, string fatherName, string number, string nationalCode, string reshte, string tavalod, string university, string imageFile)
        {
            FullName = fullName;
            FatherName = fatherName;
            Number = number;
            NationalCode = nationalCode;
            Reshte = reshte;
            Tavalod = tavalod;
            University = university;
            if (File.Exists(imageFile))
                ImageFile = Image.FromFile(imageFile);
        }
    }
}

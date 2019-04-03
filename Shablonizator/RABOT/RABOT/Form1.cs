using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
namespace RABOT
{
    public partial class Form1 : Form


        
    {
        public string myname;
        private readonly string TemplatefileName = Application.StartupPath + "\\shablon.docx";

        public Form1()
        {
            InitializeComponent();
        }
        DataTable table = new DataTable();
        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            {
                table.Columns.Add("№№ П/П", typeof(string));// data type int
                table.Columns.Add("Дата поставки", typeof(string));// datatype string
                table.Columns.Add("Наименование сопроводительного документа, №№ накладных", typeof(string));// datatype string
                table.Columns.Add("Наименование материалов и конструкции", typeof(string));// data type int
                table.Columns.Add("Кол-во", typeof(string));
                table.Columns.Add("Ед.Изм.", typeof(string));
                table.Columns.Add("Поставщик", typeof(string));
                table.Columns.Add("Отклонения от ГОСТа, СНиПа, ТУ, ВСН, Дефекты", typeof(string));
                table.Columns.Add("Подпись лица, осуществяющего конторль", typeof(string));
                table.Columns.Add("Примечание", typeof(string));



                dataGridView1.DataSource = table;
            }


            string[] lineOfContents = File.ReadAllLines("1.txt");
            foreach (var line in lineOfContents)                            // Стек из сохраненных слов combox
            {
                string[] tokens = line.Split(',');
                comboBox1.Items.Add(tokens[0]);                                
            }
                                 
            string[] lineOfContents2 = File.ReadAllLines("2.txt");
            foreach (var line in lineOfContents2)
            {
                string[] tokens2 = line.Split(',');
                comboBox2.Items.Add(tokens2[0]);
            }

            string[] lineOfContents3 = File.ReadAllLines("3.txt");
            foreach (var line in lineOfContents3)
            {
                string[] tokens3 = line.Split(',');
                comboBox3.Items.Add(tokens3[0]);
            }

            string[] lineOfContents4 = File.ReadAllLines("4.txt");
            foreach (var line in lineOfContents4)
            {
                string[] tokens4 = line.Split(',');
                comboBox4.Items.Add(tokens4[0]);
            }   
            string[] lineOfContents5 = File.ReadAllLines("5.txt");
            foreach (var line in lineOfContents5)
            {
                string[] tokens5 = line.Split(',');
                comboBox5.Items.Add(tokens5[0]);
            }

            string[] lineOfContents6 = File.ReadAllLines("6.txt");
            foreach (var line in lineOfContents6)
            {
                string[] tokens6 = line.Split(',');
                comboBox6.Items.Add(tokens6[0]);
            }

            string[] lineOfContents7 = File.ReadAllLines("7.txt");
            foreach (var line in lineOfContents7)
            {
                string[] tokens7 = line.Split(',');
                comboBox7.Items.Add(tokens7[0]);
            }

            string[] lineOfContents8 = File.ReadAllLines("8.txt");
            foreach (var line in lineOfContents8)
            {
                string[] tokens8 = line.Split(',');
                comboBox8.Items.Add(tokens8[0]);
            }

            string[] lineOfContents9 = File.ReadAllLines("9.txt");
            foreach (var line in lineOfContents9)
            {
                string[] tokens9 = line.Split(',');
                comboBox9.Items.Add(tokens9[0]);
            }

            string[] lineOfContents10 = File.ReadAllLines("10.txt");
            foreach (var line in lineOfContents10)
            {
                string[] tokens10 = line.Split(',');
                comboBox10.Items.Add(tokens10[0]);
            }

            string[] lineOfContents11 = File.ReadAllLines("11.txt");
            foreach (var line in lineOfContents11)
            {
                string[] tokens11 = line.Split(',');
                comboBox11.Items.Add(tokens11[0]);
            }

            string[] lineOfContents12 = File.ReadAllLines("12.txt");
            foreach (var line in lineOfContents12)
            {
                string[] tokens12 = line.Split(',');
                comboBox12.Items.Add(tokens12[0]);
            }

            string[] lineOfContents13 = File.ReadAllLines("13.txt");
            foreach (var line in lineOfContents13)
            {
                string[] tokens13 = line.Split(',');
                comboBox13.Items.Add(tokens13[0]);
            }

            string[] lineOfContents14 = File.ReadAllLines("14.txt");
            foreach (var line in lineOfContents14)
            {
                string[] tokens14 = line.Split(',');
                comboBox14.Items.Add(tokens14[0]);
            }

            string[] lineOfContents15 = File.ReadAllLines("15.txt");
            foreach (var line in lineOfContents15)
            {
                string[] tokens15 = line.Split(',');
                comboBox15.Items.Add(tokens15[0]);
            }

            string[] lineOfContents16 = File.ReadAllLines("16.txt");
            foreach (var line in lineOfContents16)
            {
                string[] tokens16 = line.Split(',');
                comboBox16.Items.Add(tokens16[0]);
            }

            string[] lineOfContents17 = File.ReadAllLines("17.txt");
            foreach (var line in lineOfContents17)
            {
                string[] tokens17 = line.Split(',');
                comboBox17.Items.Add(tokens17[0]);
            }

            string[] lineOfContents18 = File.ReadAllLines("18.txt");
            foreach (var line in lineOfContents18)
            {
                string[] tokens18 = line.Split(',');
                comboBox18.Items.Add(tokens18[0]);
            }

            string[] lineOfContents19 = File.ReadAllLines("19.txt");
            foreach (var line in lineOfContents19)
            {
                string[] tokens19 = line.Split(',');
                comboBox19.Items.Add(tokens19[0]);
            }

            string[] lineOfContents20 = File.ReadAllLines("20.txt");
            foreach (var line in lineOfContents20)
            {
                string[] tokens20 = line.Split(',');
                comboBox20.Items.Add(tokens20[0]);
            }

            string[] lineOfContents21 = File.ReadAllLines("21.txt");
            foreach (var line in lineOfContents21)
            {
                string[] tokens21 = line.Split(',');
                comboBox21.Items.Add(tokens21[0]);
            }

            string[] lineOfContents22 = File.ReadAllLines("22.txt");
            foreach (var line in lineOfContents22)
            {
                string[] tokens22 = line.Split(',');
                comboBox22.Items.Add(tokens22[0]);
            }

            string[] lineOfContents23 = File.ReadAllLines("23.txt");
            foreach (var line in lineOfContents23)
            {
                string[] tokens23 = line.Split(',');
                comboBox23.Items.Add(tokens23[0]);
            }

            string[] lineOfContents24 = File.ReadAllLines("24.txt");
            foreach (var line in lineOfContents24)
            {
                string[] tokens24 = line.Split(',');
                comboBox24.Items.Add(tokens24[0]);
            }

            string[] lineOfContents25 = File.ReadAllLines("25.txt");
            foreach (var line in lineOfContents25)
            {
                string[] tokens25 = line.Split(',');
                comboBox25.Items.Add(tokens25[0]);
            }

            string[] lineOfContents26 = File.ReadAllLines("26.txt");
            foreach (var line in lineOfContents26)
            {
                string[] tokens26 = line.Split(',');
                comboBox26.Items.Add(tokens26[0]);
            }

            string[] lineOfContents27 = File.ReadAllLines("27.txt");
            foreach (var line in lineOfContents27)
            {
                string[] tokens27 = line.Split(',');
                comboBox27.Items.Add(tokens27[0]);
            }

            string[] lineOfContents28 = File.ReadAllLines("28.txt");
            foreach (var line in lineOfContents28)
            {
                string[] tokens28 = line.Split(',');
                comboBox28.Items.Add(tokens28[0]);
            }

            string[] lineOfContents29 = File.ReadAllLines("29.txt");
            foreach (var line in lineOfContents29)
            {
                string[] tokens29 = line.Split(',');
                comboBox29.Items.Add(tokens29[0]);
            }

            string[] lineOfContents30 = File.ReadAllLines("30.txt");
            foreach (var line in lineOfContents30)
            {
                string[] tokens30 = line.Split(',');
                comboBox30.Items.Add(tokens30[0]);
            }







        }

        private void открытьToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            MessageBox.Show("ФАИЛ СОХРАНЕН В РАЗДЕЛ 'Документы' ");

            StreamWriter sw2 = File.AppendText("1.txt"); //Файл в который сохраняется выбранный в опен файл дайлог.
            sw2.WriteLine(comboBox1.Text);
            sw2.Close();
            StreamWriter sw3 = File.AppendText("2.txt");
            sw3.WriteLine(comboBox2.Text);
            sw3.Close();
            StreamWriter sw4 = File.AppendText("3.txt");
            sw4.WriteLine(comboBox3.Text);
            sw4.Close();
            StreamWriter sw5 = File.AppendText("4.txt");
            sw5.WriteLine(comboBox4.Text);
            sw5.Close();
            StreamWriter sw6 = File.AppendText("5.txt");
            sw6.WriteLine(comboBox5.Text);
            sw6.Close();
            StreamWriter sw7 = File.AppendText("6.txt");
            sw7.WriteLine(comboBox6.Text);
            sw7.Close();
            StreamWriter sw8 = File.AppendText("7.txt");
            sw8.WriteLine(comboBox7.Text);
            sw8.Close();
            StreamWriter sw9 = File.AppendText("8.txt");
            sw9.WriteLine(comboBox8.Text);
            sw9.Close();
            StreamWriter sw10 = File.AppendText("9.txt");
            sw10.WriteLine(comboBox9.Text);
            sw10.Close();
            StreamWriter sw11 = File.AppendText("10.txt");
            sw11.WriteLine(comboBox10.Text);
            sw11.Close();
            StreamWriter sw12 = File.AppendText("11.txt");
            sw12.WriteLine(comboBox11.Text);
            sw12.Close();
            StreamWriter sw13 = File.AppendText("12.txt");
            sw13.WriteLine(comboBox12.Text);
            sw13.Close();
            StreamWriter sw14 = File.AppendText("13.txt");
            sw14.WriteLine(comboBox13.Text);
            sw14.Close();
            StreamWriter sw15 = File.AppendText("14.txt");
            sw15.WriteLine(comboBox14.Text);
            sw15.Close();
            StreamWriter sw16 = File.AppendText("15.txt");
            sw16.WriteLine(comboBox15.Text);
            sw16.Close();
            StreamWriter sw17 = File.AppendText("16.txt");
            sw17.WriteLine(comboBox16.Text);
            sw17.Close();
            StreamWriter sw18 = File.AppendText("17.txt");
            sw18.WriteLine(comboBox17.Text);
            sw18.Close();
            StreamWriter sw19 = File.AppendText("18.txt");
            sw19.WriteLine(comboBox18.Text);
            sw19.Close();
            StreamWriter sw20 = File.AppendText("19.txt");
            sw20.WriteLine(comboBox19.Text);
            sw20.Close();
            StreamWriter sw21 = File.AppendText("20.txt");
            sw21.WriteLine(comboBox20.Text);
            sw21.Close();
            StreamWriter sw22 = File.AppendText("21.txt");
            sw22.WriteLine(comboBox21.Text);
            sw22.Close();
            StreamWriter sw23 = File.AppendText("22.txt");
            sw23.WriteLine(comboBox22.Text);
            sw23.Close();
            StreamWriter sw24 = File.AppendText("23.txt");
            sw24.WriteLine(comboBox23.Text);
            sw24.Close();
            StreamWriter sw25 = File.AppendText("24.txt");
            sw25.WriteLine(comboBox24.Text);
            sw25.Close();
            StreamWriter sw26 = File.AppendText("25.txt");
            sw26.WriteLine(comboBox25.Text);
            sw26.Close();
            StreamWriter sw27 = File.AppendText("26.txt");
            sw27.WriteLine(comboBox26.Text);
            sw27.Close();
            StreamWriter sw28 = File.AppendText("27.txt");
            sw28.WriteLine(comboBox27.Text);
            sw28.Close();
            StreamWriter sw29 = File.AppendText("28.txt");
            sw29.WriteLine(comboBox28.Text);
            sw29.Close();
            StreamWriter sw30 = File.AppendText("29.txt");
            sw30.WriteLine(comboBox29.Text);
            sw30.Close();
            StreamWriter sw31 = File.AppendText("30.txt");
            sw31.WriteLine(comboBox30.Text);
            sw31.Close();


            var n1 = comboBox1.Text;
            var n2 = comboBox2.Text;
            var n3 = comboBox3.Text;
            var n4 = comboBox4.Text;
            var n5 = comboBox5.Text;
            var n6 = comboBox6.Text;
            var n7 = comboBox7.Text;
            var n8 = comboBox8.Text;
            var n9 = comboBox9.Text;
            var n10 = comboBox10.Text;
            var n11 = comboBox11.Text;
            var n12 = comboBox12.Text;
            var n13 = comboBox13.Text;
            var n14 = comboBox14.Text;
            var n15 = comboBox15.Text;
            var n16 = comboBox16.Text;
            var n17 = comboBox17.Text;
            var n18 = comboBox18.Text;
            var n19 = comboBox19.Text;
            var n20 = comboBox20.Text;
            var n21 = comboBox21.Text;
            var n22 = comboBox22.Text;
            var n23 = comboBox23.Text;
            var n24 = comboBox24.Text;
            var n25 = comboBox25.Text;
            var n26 = comboBox26.Text;
            var n27 = comboBox27.Text;
            var n28 = comboBox28.Text;
            var n29 = comboBox29.Text;
            var n30 = comboBox30.Text;
            var vipadaet = comboBox31.Text;               //доделать, когда придёт инфа!!!
            var N1 = textBox1.Text;
            var ch = textBox2.Text;
            var chm = textBox3.Text;
            var chg = textBox4.Text;
            var zel1 = textBox5.Text;
            var zel2 = textBox6.Text;
            var zel3 = textBox7.Text;
            var zel4 = textBox8.Text;
            var zel5 = textBox9.Text;
            var zel6 = textBox10.Text;
            var zel7 = textBox11.Text;
            var datanach = textBox12.Text;
            var mes1 = textBox13.Text;
            var god1 = textBox14.Text;
            var dataokon = textBox15.Text;
            var mes2 = textBox16.Text;
            var god2 = textBox17.Text;
            var prog = textBox18.Text;
            var tp6 = textBox19.Text;
            var dopsv = textBox20.Text;
            var kolvo = textBox21.Text;
            var prilojeniya = textBox22.Text;




            comboBox1.Items.Add(n1);                     // присвоение ключей к Word file
            comboBox2.Items.Add(n2);
            comboBox3.Items.Add(n3);
            comboBox4.Items.Add(n4);
            comboBox5.Items.Add(n5);
            comboBox6.Items.Add(n6);
            comboBox7.Items.Add(n7);
            comboBox8.Items.Add(n8);
            comboBox9.Items.Add(n9);
            comboBox10.Items.Add(n10);
            comboBox11.Items.Add(n11);
            comboBox12.Items.Add(n12);
            comboBox13.Items.Add(n13);
            comboBox14.Items.Add(n14);
            comboBox15.Items.Add(n15);
            comboBox16.Items.Add(n16);
            comboBox17.Items.Add(n17);
            comboBox18.Items.Add(n18);
            comboBox19.Items.Add(n19);
            comboBox20.Items.Add(n20);
            comboBox21.Items.Add(n21);
            comboBox22.Items.Add(n22);
            comboBox23.Items.Add(n23);
            comboBox24.Items.Add(n24);
            comboBox25.Items.Add(n25);
            comboBox26.Items.Add(n26);
            comboBox27.Items.Add(n27);
            comboBox28.Items.Add(n28);
            comboBox29.Items.Add(n29);
            comboBox30.Items.Add(n30);
            comboBox31.Items.Add(vipadaet);



            // Woded Export
            var wordApp = new Word.Application();
            wordApp.Visible = false;


            try
            {
                var wordDocument = wordApp.Documents.Open(TemplatefileName);
                ReplaceWordStub("{n1}", n1, wordDocument);
                ReplaceWordStub("{n2}", n2, wordDocument);         // ключи для Word file
                ReplaceWordStub("{n3}", n3, wordDocument);
                ReplaceWordStub("{n4}", n4, wordDocument);
                ReplaceWordStub("{n5}", n5, wordDocument);
                ReplaceWordStub("{n6}", n6, wordDocument);
                ReplaceWordStub("{n7}", n7, wordDocument);
                ReplaceWordStub("{n8}", n8, wordDocument);
                ReplaceWordStub("{n9}", n9, wordDocument);
                ReplaceWordStub("{n10}", n10, wordDocument);
                ReplaceWordStub("{n11}", n11, wordDocument);
                ReplaceWordStub("{n12}", n12, wordDocument);
                ReplaceWordStub("{n13}", n13, wordDocument);
                ReplaceWordStub("{n14}", n14, wordDocument);
                ReplaceWordStub("{n15}", n15, wordDocument);
                ReplaceWordStub("{n16}", n16, wordDocument);
                ReplaceWordStub("{n17}", n17, wordDocument);
                ReplaceWordStub("{n18}", n18, wordDocument);
                ReplaceWordStub("{n19}", n19, wordDocument);
                ReplaceWordStub("{n20}", n20, wordDocument);
                ReplaceWordStub("{n21}", n21, wordDocument);
                ReplaceWordStub("{n22}", n22, wordDocument);
                ReplaceWordStub("{n23}", n23, wordDocument);
                ReplaceWordStub("{n24}", n24, wordDocument);
                ReplaceWordStub("{n25}", n25, wordDocument);
                ReplaceWordStub("{n26}", n26, wordDocument);
                ReplaceWordStub("{n27}", n27, wordDocument);
                ReplaceWordStub("{n28}", n28, wordDocument);
                ReplaceWordStub("{n29}", n29, wordDocument);
                ReplaceWordStub("{n30}", n30, wordDocument);
                ReplaceWordStub("{vipadaet}", vipadaet, wordDocument);
                ReplaceWordStub("{N1}", N1, wordDocument);
                ReplaceWordStub("{ch}",ch, wordDocument);
                ReplaceWordStub("{chm}",chm, wordDocument);
                ReplaceWordStub("{chg}", chg, wordDocument);
                ReplaceWordStub("{zel1}", zel1, wordDocument);
                ReplaceWordStub("{zel2}", zel2, wordDocument);
                ReplaceWordStub("{zel3}", zel3, wordDocument);
                ReplaceWordStub("{zel4}", zel4, wordDocument);
                ReplaceWordStub("{zel5}", zel5, wordDocument);
                ReplaceWordStub("{zel6}", zel6, wordDocument);
                ReplaceWordStub("{zel7}", zel7, wordDocument);
                ReplaceWordStub("{datanach}", datanach, wordDocument);
                ReplaceWordStub("{mes1}", mes1, wordDocument);
                ReplaceWordStub("{god1}", god1, wordDocument);
                ReplaceWordStub("{dataokon}", dataokon, wordDocument);
                ReplaceWordStub("{mes2}", mes2, wordDocument);
                ReplaceWordStub("{god2}", god2, wordDocument);
                ReplaceWordStub("{prog}", prog, wordDocument);
                ReplaceWordStub("{tp6}", tp6, wordDocument);
                ReplaceWordStub("{dopsv}", dopsv, wordDocument);
                ReplaceWordStub("{kolvo}", kolvo, wordDocument);
                ReplaceWordStub("{prilojeniya}", prilojeniya, wordDocument);



                wordDocument.SaveAs("Готовый акт.docx"); // Сохранение файла
                wordApp.Visible = true;
            }
            catch

            {
                MessageBox.Show("Произошла ошибка :) ");
            }

        }

        private void ReplaceWordStub (string stubToReplace, string text, Word.Document wordDocument)
        {
            var range = wordDocument.Content;
            range.Find.ClearFormatting();                                       // Ищет ключи в Word File
            range.Find.Execute(FindText: stubToReplace, ReplaceWith: text);




        }

        private void button3_Click(object sender, EventArgs e)
        {
            table.Rows.Add(textBox23.Text, textBox24.Text, textBox25.Text, textBox26.Text, textBox27.Text, textBox28.Text, textBox29.Text, textBox30.Text, textBox31.Text, textBox32.Text);
            dataGridView1.DataSource = table;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);
            Microsoft.Office.Interop.Excel._Worksheet worksheet = null;
            worksheet = workbook.Sheets["Лист1"];
            worksheet = workbook.ActiveSheet;
            worksheet.Name = "Таблица";

            for (int i = 1; i < dataGridView1.Columns.Count + 1; i++)
            {
                worksheet.Cells[1, i] = dataGridView1.Columns[i - 1].HeaderText;

            }

            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView1.Columns.Count; j++)
                {
                    worksheet.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value;
                }
            }

            var saveFileDialoge = new SaveFileDialog();
            saveFileDialoge.FileName = "Материалы";
            saveFileDialoge.DefaultExt = "xlsx";
            if (saveFileDialoge.ShowDialog() == DialogResult.OK)
            {
                workbook.SaveAs(saveFileDialoge.FileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            }
            app.Quit();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            File.WriteAllText(@"1.txt", string.Empty);
            File.WriteAllText(@"2.txt", string.Empty); //удалает все сохраненые слова
            File.WriteAllText(@"3.txt", string.Empty);
            File.WriteAllText(@"4.txt", string.Empty);
            File.WriteAllText(@"5.txt", string.Empty);
            File.WriteAllText(@"6.txt", string.Empty);
            File.WriteAllText(@"7.txt", string.Empty);
            File.WriteAllText(@"8.txt", string.Empty);
            File.WriteAllText(@"9.txt", string.Empty);
            File.WriteAllText(@"10.txt", string.Empty);
            File.WriteAllText(@"11.txt", string.Empty);
            File.WriteAllText(@"12.txt", string.Empty);
            File.WriteAllText(@"13.txt", string.Empty);
            File.WriteAllText(@"14.txt", string.Empty);
            File.WriteAllText(@"15.txt", string.Empty);
            File.WriteAllText(@"16.txt", string.Empty);
            File.WriteAllText(@"17.txt", string.Empty);
            File.WriteAllText(@"18.txt", string.Empty);
            File.WriteAllText(@"19.txt", string.Empty);
            File.WriteAllText(@"20.txt", string.Empty);
            File.WriteAllText(@"21.txt", string.Empty);
            File.WriteAllText(@"22.txt", string.Empty);
            File.WriteAllText(@"23.txt", string.Empty);
            File.WriteAllText(@"24.txt", string.Empty);
            File.WriteAllText(@"25.txt", string.Empty);
            File.WriteAllText(@"26.txt", string.Empty);
            File.WriteAllText(@"27.txt", string.Empty);
            File.WriteAllText(@"28.txt", string.Empty);
            File.WriteAllText(@"29.txt", string.Empty);
            File.WriteAllText(@"30.txt", string.Empty);


        }



        private void label1_Click_1(object sender, EventArgs e)
        {

        }

        
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void label10_Click(object sender, EventArgs e)
        {

        }

        private void label12_Click(object sender, EventArgs e)
        {

        }

        private void label23_Click(object sender, EventArgs e)
        {

        }

        private void label22_Click(object sender, EventArgs e)
        {

        }

        private void label21_Click(object sender, EventArgs e)
        {

        }

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void comboBox20_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void label24_Click(object sender, EventArgs e)
        {

        }

        private void comboBox19_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox23_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox21_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void label27_Click(object sender, EventArgs e)
        {

        }

        private void label13_Click(object sender, EventArgs e)
        {

        }

        private void label16_Click(object sender, EventArgs e)
        {

        }

        private void label17_Click(object sender, EventArgs e)
        {

        }

        private void comboBox13_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void label21_Click_1(object sender, EventArgs e)
        {

        }

        private void label24_Click_1(object sender, EventArgs e)
        {

        }

        private void comboBox25_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void label40_Click(object sender, EventArgs e)
        {

        }

        private void comboBox27_SelectedIndexChanged(object sender, EventArgs e)
        {

        }        

        private void comboBox30_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox22_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox23_SelectedIndexChanged_1(object sender, EventArgs e)
        {

        }

        private void label53_Click(object sender, EventArgs e)
        {

        }

        private void comboBox29_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void label59_Click(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }

        private void label72_Click(object sender, EventArgs e)
        {

        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox8_TextChanged(object sender, EventArgs e)
        {

        }

        private void label69_Click(object sender, EventArgs e)
        {

        }

        private void label78_Click(object sender, EventArgs e)
        {

        }

        private void comboBox31_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void label85_Click(object sender, EventArgs e)
        {

        }

        private void label86_Click(object sender, EventArgs e)
        {

        }

        private void label88_Click(object sender, EventArgs e)
        {

        }

        private void label91_Click(object sender, EventArgs e)
        {

        }

        private void textBox32_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox23_TextChanged(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void tabPage4_Click(object sender, EventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {
            tabControl2.Visible = false;
            button5.Visible = false;


        }

        private void tabPage4_Click_1(object sender, EventArgs e)
        {

        }

        private void button6_Click(object sender, EventArgs e)
        {

        }

        

        private void button8_Click_1(object sender, EventArgs e)
        {
            tabControl2.Visible = true;
            button5.Visible = true;
        }
    }
}

using Microsoft.VisualBasic.FileIO;
using System;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;
using System.Collections.Generic;
using System.Text;


namespace Test
{
    

    public partial class Form1 : Form
    {
        private OleDbConnection con;
        public Form1()
        {
            InitializeComponent();

        }

        OpenFileDialog fd = new OpenFileDialog();
        private void button1_Click(object sender, EventArgs e)
        {
            //Выгружаем из файла в datagridview
            fd.Filter = "CSV|*.csv";
            if (fd.ShowDialog() == DialogResult.OK)
            {
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                textBox1.Text = fd.FileName;
                var reader = new StreamReader(@fd.FileName, Encoding.GetEncoding(1251));
                

                    List<string> listA = new List<string>();
                    List<string> listB = new List<string>();
                    List<string> listC = new List<string>();
                    while (!reader.EndOfStream)
                    {
                        var line = reader.ReadLine();
                        var values = line.Split(';');
                        try
                        {
                            listA.Add(values[0]);
                            listB.Add(values[1]);
                            listC.Add(values[2]);
                        }
                        catch (Exception)
                        {

                        }
                    }
                    //Заполняем заголовки
                    dataGridView1.Columns[0].HeaderText = listA[2];
                    dataGridView1.Columns[1].HeaderText = listB[2];
                    dataGridView1.Columns[2].HeaderText = listB[3];
                    dataGridView1.Columns[3].HeaderText = listB[4];
                    dataGridView1.Columns[4].HeaderText = listC[2];
                string main="";
                string second="";
                //Заполнаем датагрид
                for (int i = 5, j = 0; i < listC.Count; i++, j++)
                {
                    dataGridView1.Rows.Add();
                    int count12 = 0;
                    foreach (char c in listA[i])
                        if (c == '.') count12++;
                    if (listA[i] != "")
                    {
                        switch (count12)
                        {
                            case 0:
                                main = listB[i];
                                dataGridView1[0, j].Value = listA[i];
                                dataGridView1[1, j].Value = main;
                                break;
                            case 1:
                                second = listB[i];
                                dataGridView1[0, j].Value = listA[i];
                                dataGridView1[1, j].Value = main;
                                dataGridView1[2, j].Value = second;
                                break;
                            case 2:
                                dataGridView1[0, j].Value = listA[i];
                                dataGridView1[1, j].Value = main;
                                dataGridView1[2, j].Value = second;
                                dataGridView1[3, j].Value = listB[i];
                                dataGridView1[4, j].Value = listC[i];
                                break;
                        }
                    }
                    else
                    {
                        dataGridView1[0, j].Value = listA[i];
                        dataGridView1[4, j].Value = listC[i];
                    }
                }



            }
        }

    

     
        private void button2_Click(object sender, EventArgs e)
        {
            //Открываем бд и линкуем 
            fd.Filter = "ACCDB |*.accdb";
            if (fd.ShowDialog() == DialogResult.OK)
            {
                textBox2.Text = fd.FileName;
                con = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source = " +@fd.FileName);
                con.Open();
               

            }
        }

        private void button3_Click(object sender, EventArgs e)
        {   //Этот огромный кусок кода ловит ошибки
            if (textBox1.Text == "")
            {
                textBox3.Clear();
                textBox3.Text = "Файл csv не выбран";
            }
            else
            {
                //Считываем из datagridview в бд
                string listBB = "";
                string listC1 = "";
                string listC2 = "";
                string listC3 = "";
                string help1 = "";
                string help2 = "";
                string a = "";
                string b = "";
                int f1 = 1;
                int f2 = 1;
                int f3 = 1;
                int f4 = 1;
                try
                {
                    for (int i = 0; i < dataGridView1.RowCount; i++)
                    {
                        if (dataGridView1[1, i].Value != null && dataGridView1[1, i].Value.ToString().Length != 0 && CountWords(listBB, dataGridView1[1, i].Value.ToString()) == 0)
                        {
                            //Заполняем таблицы уникальными значениями
                            listBB += dataGridView1[1, i].Value.ToString();
                            a = dataGridView1[0, i].Value.ToString();
                            b = dataGridView1[1, i].Value.ToString();
                            help2 = b;
                            string query = "INSERT INTO [Группы основные] ([Код процесса], [Наименование процесса]) VALUES ('" + f1 + "','" + b + "')";
                            OleDbCommand command = new OleDbCommand(query, con);
                            command.ExecuteNonQuery();
                            f1++;
                        }
                        if (dataGridView1[2, i].Value != null && dataGridView1[2, i].Value.ToString().Length != 0 && CountWords(listC1, dataGridView1[2, i].Value.ToString()) == 0)
                        {
                            listC1 += dataGridView1[2, i].Value.ToString();
                            a = dataGridView1[0, i].Value.ToString();
                            b = dataGridView1[2, i].Value.ToString();
                            help1 = b;
                            string query = "INSERT INTO [Группы вторичные] ([Код репозитория], [Репозиторий БП],[Наименование процесса]) VALUES ('" + f2 + "','" + b + "','" + help2 + "')";
                            OleDbCommand command = new OleDbCommand(query, con);
                            command.ExecuteNonQuery();
                            f2++;
                        }
                        if (dataGridView1[3, i].Value != null && dataGridView1[3, i].Value.ToString().Length != 0 && CountWords(listC2, dataGridView1[3, i].Value.ToString()) == 0)
                        {
                            listC2 += dataGridView1[3, i].Value.ToString();
                            a = dataGridView1[0, i].Value.ToString();
                            b = dataGridView1[3, i].Value.ToString();
                            string query = "INSERT INTO [Группы третичные] ([Код процесса обслуживания], [Процессы обслуживания ФЛ],[Репозиторий БП]) VALUES ('" + f3 + "','" + b + "','" + help1 + "')";
                            OleDbCommand command = new OleDbCommand(query, con);
                            command.ExecuteNonQuery();
                            f3++;
                        }
                        if (dataGridView1[4, i].Value != null && dataGridView1[4, i].Value.ToString().Length != 0 && CountWords(listC3, dataGridView1[4, i].Value.ToString()) == 0)
                        {
                            listC3 += dataGridView1[4, i].Value.ToString();
                            b = dataGridView1[4, i].Value.ToString();
                            string query = "INSERT INTO [Владельцы процесса] ([Владелец],[Код владельца]) VALUES ('" + b + "','" + f4 + "')";
                            OleDbCommand command = new OleDbCommand(query, con);
                            command.ExecuteNonQuery();
                            f4++;
                        }

                    }
                    textBox3.Clear();
                    textBox3.Text = "OK";
                }

                catch (System.InvalidOperationException)
                {
                    textBox3.Clear();
                    textBox3.Text = "База данных не выбрана";
                }
                catch (System.Data.OleDb.OleDbException)
                {
                    textBox3.Clear();
                    textBox3.Text = "База данных не пустая";
                }
                
            }
            
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            dataGridView1.ColumnCount = 5;
            dataGridView1.RowCount = 1;
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
        private int CountWords(string s, string s0)
        {
                int count = (s.Length - s.Replace(s0, "").Length) / s0.Length;
                return count;
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (con!=null)
            con.Close();
        }

        private void button4_Click(object sender, EventArgs e)
        {
           
        }

        private void button4_Click_1(object sender, EventArgs e)
        {

        }
    }
}

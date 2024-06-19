using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace Lab26
{
    public partial class Form1 : Form
    {
        private Dictionary<string, string> templates = new Dictionary<string, string>()
        {
            { "������ 1", "������1.docx" },
            { "������ 2", "������2.docx" }
        };

        private string newDocumentPath;

        public Form1()
        {
            InitializeComponent();

            comboBoxTemplates.DataSource = new BindingSource(templates, null);
            comboBoxTemplates.DisplayMember = "Key";
            comboBoxTemplates.ValueMember = "Value";
        }

        // ������ ��� ��������� ��������� � �������
        private void button1_Click(object sender, EventArgs e)
        {
            string selectedTemplate = ((KeyValuePair<string, string>)comboBoxTemplates.SelectedItem).Value;
            var helper = new WordHelper(selectedTemplate);

            var items = new Dictionary<string, string>
            {
                {"[����]", textBox1.Text },
                {"[ϲ� ����������]", textBox2.Text },
                {"[������ ����������]", textBox3.Text },
                {"[������� ����������]", textBox4.Text },
                {"[������, ������� ����������]", textBox5.Text },
                {"[̳��� ����������]", textBox7.Text },
                {"[������� ����������]", textBox8.Text },
                {"[�������� ������ ����������]", textBox9.Text },

                {"[ϲ� ����������]", textBox10.Text },
                {"[������ ����������]", textBox12.Text },
                {"[������� ����������]", textBox11.Text },
                {"[������, ������� ����������]", textBox14.Text },
                {"[̳��� ����������]", textBox16.Text },
                {"[������� ����������]", textBox15.Text },
                {"[�������� ������ ����������]", textBox17.Text },
            };

            newDocumentPath = helper.Process(items);
            if (!string.IsNullOrEmpty(newDocumentPath))
            {
                MessageBox.Show("�������� ������ �������� � ���������.");
            }
            else
            {
                MessageBox.Show("������� ��� �������� ���������.");
            }
        }

        // ������, �� ���� �� ������ ������ ��������
        private void button3_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(newDocumentPath))
            {
                MessageBox.Show("�������� ������� ��������.");
                return;
            }

            var helper = new WordHelper(newDocumentPath);

            bool result = helper.FindAndReplace(textBoxFind.Text, textBoxReplace.Text);
            if (result)
            {
                MessageBox.Show("����� � ����� ������� ������.");
            }
            else
            {
                MessageBox.Show("������� ��� �������� ������ � �����.");
            }
        }
    }
}

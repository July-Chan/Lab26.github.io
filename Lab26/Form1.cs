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
            { "Шаблон 1", "Шаблон1.docx" },
            { "Шаблон 2", "Шаблон2.docx" }
        };

        private string newDocumentPath;

        public Form1()
        {
            InitializeComponent();

            comboBoxTemplates.DataSource = new BindingSource(templates, null);
            comboBoxTemplates.DisplayMember = "Key";
            comboBoxTemplates.ValueMember = "Value";
        }

        // Кнопка для створення документу з шаблону
        private void button1_Click(object sender, EventArgs e)
        {
            string selectedTemplate = ((KeyValuePair<string, string>)comboBoxTemplates.SelectedItem).Value;
            var helper = new WordHelper(selectedTemplate);

            var items = new Dictionary<string, string>
            {
                {"[Дата]", textBox1.Text },
                {"[ПІБ відправника]", textBox2.Text },
                {"[Посада відправника]", textBox3.Text },
                {"[Компанія відправника]", textBox4.Text },
                {"[Вулиця, будинок відправника]", textBox5.Text },
                {"[Місто відправника]", textBox7.Text },
                {"[Область відправника]", textBox8.Text },
                {"[Поштовий індекс відправника]", textBox9.Text },

                {"[ПІБ одержувача]", textBox10.Text },
                {"[Посада одержувача]", textBox12.Text },
                {"[Компанія одержувача]", textBox11.Text },
                {"[Вулиця, будинок одержувача]", textBox14.Text },
                {"[Місто одержувача]", textBox16.Text },
                {"[Область одержувача]", textBox15.Text },
                {"[Поштовий індекс одержувача]", textBox17.Text },
            };

            newDocumentPath = helper.Process(items);
            if (!string.IsNullOrEmpty(newDocumentPath))
            {
                MessageBox.Show("Документ успішно створено і збережено.");
            }
            else
            {
                MessageBox.Show("Помилка при створенні документу.");
            }
        }

        // Кнопка, що шукає та заміняє введені значення
        private void button3_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(newDocumentPath))
            {
                MessageBox.Show("Спочатку створіть документ.");
                return;
            }

            var helper = new WordHelper(newDocumentPath);

            bool result = helper.FindAndReplace(textBoxFind.Text, textBoxReplace.Text);
            if (result)
            {
                MessageBox.Show("Пошук і заміна виконані успішно.");
            }
            else
            {
                MessageBox.Show("Помилка при виконанні пошуку і заміни.");
            }
        }
    }
}

using GroupDocs.Editor.Formats;
using GroupDocs.Editor.HtmlCss.Resources;
using GroupDocs.Editor.Options;
using GroupDocs.Editor;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using System.CodeDom.Compiler;
using System.Diagnostics;
using Xceed.Words.NET;
using Xceed.Document.NET;
using static Recibos.ConverterNumeros;
using System.Globalization;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Text.RegularExpressions;
using Microsoft.Office;
using System.Threading;
using System.Security.Cryptography;

namespace Recibos
{
    public partial class Form1 : Form
    {
        string Nome = string.Empty, Valor = string.Empty, Servico = string.Empty, Data = string.Empty, Assinatura = string.Empty;
        string VisualizacaoText = string.Empty;

        string ReciboCaminhoFinal = string.Empty;
        public Form1()
        {
            InitializeComponent();
        }
        private bool SalvarRecibo(string nome, string valor, string servico, string data, string assinatura, string caminho, bool imprimir)
        {
            if (string.IsNullOrEmpty(nome))
                return false;

            if (string.IsNullOrEmpty(valor))
                valor = "0";

            if (string.IsNullOrEmpty(data))
                data = dateTimePicker1.Text;

            if (string.IsNullOrEmpty(assinatura))
                assinatura = "RURAL PLAN PLANEJAMENTO E CONSULTORIA";

            if (string.IsNullOrEmpty(caminho))
                caminho = Environment.CurrentDirectory;

            try
            {
                DirectoryInfo di = new DirectoryInfo(Environment.CurrentDirectory + "\\RECIBO_ORIGINAL.docx");

                using (DocX document = DocX.Load(di.FullName))
                {
                    ReciboCaminhoFinal = String.Format(caminho + "\\RECIBO_{0}.docx", Nome);
                    document.ReplaceText(new StringReplaceTextOptions() { SearchValue = "#NOME_DO_CLIENTE", NewValue = nome });
                    document.ReplaceText(new StringReplaceTextOptions() { SearchValue = "#VALOR ", NewValue = valor });
                    document.ReplaceText(new StringReplaceTextOptions() { SearchValue = "#VALOR_POR_EXTENSO", NewValue = toExtenso(decimal.Parse(valor, CultureInfo.InvariantCulture)) });
                    document.ReplaceText(new StringReplaceTextOptions() { SearchValue = "#SERVICE ", NewValue = servico });
                    document.ReplaceText(new StringReplaceTextOptions() { SearchValue = "#DATA_DOC   ", NewValue = data });
                    document.ReplaceText(new StringReplaceTextOptions() { SearchValue = "#ASSINATURA ", NewValue = assinatura });
                    document.SaveAs(ReciboCaminhoFinal);
                    if (imprimir)
                    {
                        Microsoft.Office.Interop.Word.Application ap = new Microsoft.Office.Interop.Word.Application();
                        Microsoft.Office.Interop.Word.Document documentp = ap.Documents.Open(ReciboCaminhoFinal);

                        documentp.PageSetup.TopMargin = 50;
                        documentp.PageSetup.RightMargin = 50;
                        documentp.PageSetup.BottomMargin = 50;
                        documentp.PageSetup.LeftMargin = 50;
                        documentp.PageSetup.PaperSize = Microsoft.Office.Interop.Word.WdPaperSize.wdPaperA4;

                        Microsoft.Office.Interop.Word.WdPrintOutRange printRange = Microsoft.Office.Interop.Word.WdPrintOutRange.wdPrintCurrentPage;

                        documentp.PrintOut(false, null, printRange);
                        documentp.Close(false, false, false);

                        int count = 0;
                        while (count < 100)
                        {
                            
                            MostrarPainelProgresso(count);
                            count++;

                            Thread.Sleep(50);
                        }

                        FileInfo arquivo = new FileInfo(ReciboCaminhoFinal);
                        arquivo.Delete();
                    }
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Não foi possível emitir o Recibo. ", "Erro ao emitir Recibo");
                //throw;

                return false;
            }
            
            return true;
        }
        public void SetarVisualizacao()
        {
            string valor_extenso;

            if (string.IsNullOrEmpty(Valor))
                valor_extenso = "Zero Reais";
            else valor_extenso = toExtenso(decimal.Parse(Valor, CultureInfo.InvariantCulture));

            if (string.IsNullOrEmpty(Assinatura))
                Assinatura = "RURAL PLAN PLANEJAMENTO E CONSULTORIA";

            string formated_text = string.Format(VisualizacaoText, Nome, Valor, valor_extenso, Servico, dateTimePicker1.Text, Assinatura);
            richTextBox1.Text = formated_text;
        }
        public void LimparCampos()
        {
            Nome = string.Empty;
            Valor = string.Empty;
            Servico = string.Empty;
            Assinatura = string.Empty;

            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "RURAL PLAN PLANEJAMENTO E CONSULTORIA";

            SetarVisualizacao();
        }
        public void MostrarPainelProgresso(int valorbarra)
        {
            if (!panel3.Enabled)
            {
                panel3.Enabled = true;
                panel3.Visible = true;
            }

            progressBar1.Value = valorbarra;
            label7.Text = valorbarra.ToString();
            if(valorbarra >= 99)
            {
                panel3.Enabled = false;
                panel3.Visible = false;
            }
        }
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            Nome = textBox1.Text;
            SetarVisualizacao();
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            Regex extractNumberRegex = new Regex("(?:-(?:[1-9](?:\\d{0,2}(?:,\\d{3})+|\\d*))|(?:0|(?:[1-9](?:\\d{0,2}(?:,\\d{3})+|\\d*))))(?:.\\d+|)");

            string[] extracted = extractNumberRegex.Matches(textBox2.Text)
            .Cast<Match>()
            .Select(m => m.Value)
            .ToArray();
            Valor = string.Concat(extracted);

            SetarVisualizacao();
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            Servico = textBox3.Text;
            SetarVisualizacao();
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            Data = dateTimePicker1.Text;
            SetarVisualizacao();
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            Assinatura = textBox4.Text;
            SetarVisualizacao();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            //Apagar

            LimparCampos();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //Salvar arquivo

            //SalvarRecibo(Nome, Valor, Servico, Data, Assinatura, "");

            using (var fbd = new FolderBrowserDialog())
            {
                fbd.ShowNewFolderButton = true;
                DialogResult result = fbd.ShowDialog();

                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    if(SalvarRecibo(Nome, Valor, Servico, Data, Assinatura, fbd.SelectedPath, false))
                        MessageBox.Show("O Recibo RECIBO_" + Nome + " foi salvo em: " + fbd.SelectedPath, "Message");

                    LimparCampos();
                }
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            string nomeImpressoraPadrao = (new PrinterSettings()).PrinterName;

            if (string.IsNullOrEmpty(nomeImpressoraPadrao) || nomeImpressoraPadrao == "Microsoft Print to PDF")
            {
                MessageBox.Show("Não foi possível imprimir o Recibo.\nO sistema não encontrou nenhuma impressora padrão.", "Nenhuma impressora");
                return;
            }
            SalvarRecibo(Nome, Valor, Servico, Data, Assinatura, "", true);
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            VisualizacaoText = richTextBox1.Text;
            richTextBox1.ReadOnly = true;
            SetarVisualizacao();

            panel3.Enabled = false;
            panel3.Visible = false;

            string nomeImpressoraPadrao = (new PrinterSettings()).PrinterName;

            if (nomeImpressoraPadrao == "Microsoft Print to PDF")
                nomeImpressoraPadrao = "Nenhuma impressora encontrada.";

            label7.Text = nomeImpressoraPadrao;
        }

        //Controles não usados
        private void label2_Click(object sender, EventArgs e)
        {
            
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }
        private void label5_Click(object sender, EventArgs e)
        {

        }
        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void richTextBox1_TextChanged_1(object sender, EventArgs e)
        {

        }
    }
}

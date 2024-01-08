using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO.Ports;
using Excel = Microsoft.Office.Interop.Excel;

namespace NFC
{
    public partial class Form1 : Form
    {
        private SerialPort serialPort;
        private Excel.Application excelApp;
        private Excel.Workbook workbook;
        private Excel.Worksheet worksheet;
        private int rowNum = 1;
        public Form1()
        {
            InitializeComponent();
            InitializeSerialPort();
            InitializeExcel();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void InitializeExcel()
        {
            try 
            {
                //Inicia o Excel
                excelApp = new Excel.Application();
                excelApp.Visible = true;

                //Adiciona um novo workbook e worksheet 
                workbook =  excelApp.Workbooks.Add();
                worksheet = (Excel.Worksheet)workbook.Worksheets[1];

                //Adiciona cabe~çalhos á primeira linha da planilha 
                worksheet.Cells[1, 1] = "Tag RFID";
            }
            catch(Exception ex) 
            {
                MessageBox.Show($"Erro ao inicializar o Excel: {ex.Message}");
            }
        }

        private void InitializeSerialPort()
        {
            serialPort = new SerialPort("COM4", 9600); // Substitua "COMx" pela porta COM do seu Arduino
            serialPort.DataReceived += SerialPort_DataReceived;

            try
            {
                serialPort.Open();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao abrir a porta serial: {ex.Message}");
            }
        }

        private void SerialPort_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            if (serialPort.IsOpen)
            {
                try
                {
                    string dataReceived = serialPort.ReadLine();
                    Invoke(new Action(() =>
                    {
                        textBox1.AppendText(dataReceived + Environment.NewLine);
                        //Adiciona os dados ao Excel
                        worksheet.Cells[++rowNum, 1] = dataReceived;
                    }));
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Erro ao ler dados: {ex.Message}");
                }

                try
                {
                    string dataReceived = serialPort.ReadLine();
                    // Atualize a interface gráfica com os dados recebidos
                    Invoke(new Action(() => textBox1.AppendText(dataReceived + Environment.NewLine))); 
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Erro ao ler dados: {ex.Message}");
                }
            }
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (serialPort.IsOpen)
            {
                serialPort.Close();
            }

            //Salva e fecha o workbook ao fechar o formulário 
            if(workbook != null) 
            {
                workbook.SaveAs("Caminho\\Para\\O\\Arquivo\\Excel.xlsx"); //Dentro dos parenteses voce cola o caminho onde é pra salvar o excel, o exemplo ta escrito ali dentr    o 
                workbook.Close();
            }

            excelApp.Quit();
        }
    }
}

//C# İLE KODLANDI
using System;
using System.IO.Ports;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Collections.Generic;
using System.Data;


namespace GroundStation
{
    public partial class Form1 : Form
    {
        private List<string> columnHeader = new List<string>
    {
        "PaketNo", "Irtifa", "Enlem", "Boylam", "InisHizi", "Bilgi",
        "Sicaklik", "Basinc", "Nem", "IntubePressure1", "IntubePressure2"
    };
        private bool dataFinished = false; // Veriler bitti mi?
        public Form1()
        {
            InitializeComponent(); //Windows Form'un bileşenlerini başlatan metod.
            InitializeTimer();//timer başlatmak için metod çağırdık.
            barAvoyonik.Value = 0;//barAvoyoniğe 0 değrini atadık.
        }
        private int currentRowIndex = 0; // Başlangıçta 0. satırdan başla
        private List<string[]> data = new List<string[]>(); // Verileri bir dizi listesinde sakla

        private void InitializeTimer()
        {
            Timer timer = new Timer(); //timer nesnesini oluşturduk.
            timer.Interval = 1000; // 1 saniye
            timer.Tick += new EventHandler(timer1_Tick);
            timer.Start(); //timer ı başlattık.
        }

        OleDbConnection con;//excel ile iletişim kurduk
        OleDbCommand cmd;//excel sorgularını çalıştırmak için kullandık.
        OleDbDataReader dr;//exceldeki verileri okumak için çağırdık.
        private void Form1_Load(object sender, EventArgs e)
        {
            timer1.Start();//form yüklendiğinde timer başlattık.

            // Excel bağlantısını yaptık dosya yoluyla
            OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\DELL\Desktop\Data.xlsx;Extended Properties='Excel 12.0 Xml;HDR=YES';");
            // Sayfa1'deki verileri çektik
            OleDbCommand cmd = new OleDbCommand("Select * From [Sayfa1$]", con);

            // Bağlantıyı açık tuttuk
            con.Open();

            OleDbDataReader dr = cmd.ExecuteReader();//veritabanı sorgusunu çalıştırdık ve sorgu sonucundaki verileri OleDbDataReader nesnesi üzerinden okumak için bir veri akışı (data stream) oluşturduk. 

            // Verileri döngü içine aldık ve çağırdık.
            while (dr.Read())
            {
                string[] rowData = new string[]
                {
                    dr["PaketNo"].ToString(),
                    dr["Irtifa"].ToString(),
                    dr["Enlem"].ToString(),
                    dr["Boylam"].ToString(),
                    dr["InisHizi"].ToString(),
                    dr["Bilgi"].ToString(),
                    dr["Sicaklik"].ToString(),
                    dr["Basinc"].ToString(),
                    dr["Nem"].ToString(),
                    dr["IntubePressure1"].ToString(),
                    dr["IntubePressure2"].ToString()
            };

                data.Add(rowData);

            }
            // Verileri aldıktan sonra bağlantıyı kapattık
            con.Close();

            //seri portları aldık
            string[] portlar = SerialPort.GetPortNames();

            // ComboBox içine seri portları ekledik
            foreach (string port in portlar)
            {
                //combobox lardaki sanal portlarımızı gösterdik burada 
                comboBoxPort.Items.Add(port);
                comboBoxPort2.Items.Add(port);
                comboBoxPort3.Items.Add(port);

            }
            //comboBox içindeki değerlerden uygulama çalıştığında en baştaki port seçili hale getirdik.
            comboBoxPort.SelectedIndex = 0;
            comboBoxPort2.SelectedIndex = 0;
            comboBoxPort3.SelectedIndex = 0;
            comboBoxBaud.SelectedIndex = 0;
            comboBoxBaud2.SelectedIndex = 0;
            comboBoxBaud3.SelectedIndex = 0;
        }

        private void btnBaglan_Click(object sender, EventArgs e)
        {

            serialPort1.PortName = comboBoxPort.Text;
            serialPort1.BaudRate = 9600;
            serialPort1.Parity = Parity.Even;
            serialPort1.StopBits = StopBits.One;
            serialPort1.DataBits = 8;
            //karşılaşılan hataya rağmen programımız devam etmesi için try catch içine alacağız seriaalPort1.Open(); ı
            try
            {
                serialPort1.Open();
            }
            catch (Exception ex)
            {
                // Bağlantı başarısızsa mesaj ekle
                string hataMesaji = "Ana Avoyonik Bağlantı başarısız oldu." + ex.Message; ;
                textBoxMessage.AppendText(hataMesaji + Environment.NewLine);
            }
            //Errordan sonra düzgün bağlandığımızda o zamna sorun olmaması için serialPort1 açıkken 
            if (serialPort1.IsOpen)
            {
                // Mevcut seri portu kapat
                serialPort1.Close();
                // Bağlantı başarılıysa, düğmeyi tekrar etkinleştirin.
                btnBaglan.Enabled = false;

                // Bağlantı başarılıysa mesajı oluşturun ve TextBox'a ekleyin
                DateTime connectTime = DateTime.Now;
                string message = connectTime.ToString("dd/MM/yyyy HH:mm:ss") + " Ana Avoyonik Bağlantı";
                textBoxMessage.AppendText(message + Environment.NewLine);
            }

        }

        private void btnBaglan2_Click(object sender, EventArgs e)
        {
            serialPort1.PortName = comboBoxPort2.Text;
            serialPort1.BaudRate = 9600;
            serialPort1.Parity = Parity.Even;
            serialPort1.StopBits = StopBits.One;
            serialPort1.DataBits = 8;

            //karşılaşılan hataya rağmen programımız devam etmesi için try catch içine alacağız seriaalPort1.Open(); ı
            try
            {
                serialPort1.Open();
            }
            catch (Exception ex)
            {
                // Bağlantı başarısızsa mesaj ekle
                string hataMesaji = "Payload Bağlantı başarısız oldu." + ex.Message; ;
                textBoxMessage.AppendText(hataMesaji + Environment.NewLine);

            }
            //Errordan sonra düzgün bağlandığımızda o zamna sorun olmaması için serialPort1 açıkken 
            if (serialPort1.IsOpen)
            {

                // Mevcut seri portu kapat
                serialPort1.Close();
                // Bağlantı başarılıysa, düğmeyi tekrar etkinleştirin.
                btnBaglan2.Enabled = false;

                // Bağlantı başarılıysa mesajı oluşturun ve TextBox'a ekleyin
                DateTime connectTime = DateTime.Now;
                string message = connectTime.ToString("dd/MM/yyyy HH:mm:ss") + " Payload Bağlantı";
                textBoxMessage.AppendText(message + Environment.NewLine);
            }
        }

        private void btnBaglan3_Click(object sender, EventArgs e)
        {
            serialPort1.PortName = comboBoxPort3.Text;
            serialPort1.BaudRate = 9600;
            serialPort1.Parity = Parity.Even;
            serialPort1.StopBits = StopBits.One;
            serialPort1.DataBits = 8;

            //karşılaşılan hataya rağmen programımız devam etmesi için try catch içine alacağız seriaalPort1.Open(); ı
            try
            {
                serialPort1.Open();
            }
            catch (Exception ex)
            {
                // Bağlantı başarısızsa mesaj ekle
                string hataMesaji = "HYI Bağlantı başarısız oldu." + ex.Message; ;
                textBoxMessage.AppendText(hataMesaji + Environment.NewLine);

            }
            //Errordan sonra düzgün bağlandığımızda o zamna sorun olmaması için serialPort1 açıkken 
            if (serialPort1.IsOpen)
            {
                // Mevcut seri portu kapat
                serialPort1.Close();
                // Bağlantı başarılıysa, düğmeyi tekrar etkinleştirin.
                btnBaglan3.Enabled = false;

                // Bağlantı başarılıysa mesajı oluşturun ve TextBox'a ekleyin
                DateTime connectTime = DateTime.Now;
                string message = connectTime.ToString("dd/MM/yyyy HH:mm:ss") + " HYI Bağlantı";
                textBoxMessage.AppendText(message + Environment.NewLine);
            }
        }

        private void btnYenile_Click(object sender, EventArgs e)
        {

            // Mevcut seri portu kapat
            if (serialPort1.IsOpen)
            {
                serialPort1.Close();
            }

            // Yeni seri portu ayarla
            serialPort1.PortName = comboBoxPort3.Text;
            serialPort1.BaudRate = 9600;
            serialPort1.Parity = Parity.Even;
            serialPort1.StopBits = StopBits.One;
            serialPort1.DataBits = 8;

            serialPort1.Open();

            //// Verileri temizle
            //data.Clear();

            // Bağlantı işlemi tamamlandığında bir uyarı mesajı görüntüle
            DateTime connectTime = DateTime.Now;
            string message = connectTime.ToString("dd/MM/yyyy HH:mm:ss") + "Portlar Yenilendi";
            textBoxMessage.AppendText(message + Environment.NewLine);
        }
        private void timer1_Tick(object sender, EventArgs e)
        {
            if (dataFinished)
            {
                ((System.Timers.Timer)sender).Stop();
                return;
            }
            if (data.Count == 0)
            {
                return; // Veri yoksa işlemi durdur
            }

            if (currentRowIndex >= data.Count)
            {
                currentRowIndex = 0; // Tüm satırlar işlendiyse sıfırla
            }

            string[] rowData = data[currentRowIndex];

            for (int i = 0; i < columnHeader.Count && i < rowData.Length; i++)
            {
                string columnHeaderText = columnHeader[i];
                string columnValue = rowData[i];
                string labelName = "label" + columnHeaderText;
                Control[] matches = this.Controls.Find(labelName, true);

                if (matches.Length > 0 && matches[0] is Label)
                {
                    Label label = (Label)matches[0];
                    label.Text = columnHeaderText + ": " + columnValue;
                }
            }

            if (double.TryParse(rowData[1], out double irtifaDouble) &&
                double.TryParse(rowData[6], out double sicaklikDouble) &&
                double.TryParse(rowData[4], out double inisHiziDouble) &&
                double.TryParse(rowData[8], out double nemDouble) &&
                double.TryParse(rowData[7], out double basincDouble))
            {
                // Irtifa grafiğine veriyi ekleyin
                chart1.Series["Irtifa"].Points.Add(irtifaDouble);
                // Sıcaklık grafiğine veriyi ekleyin
                chart2.Series["Sicaklik"].Points.Add(sicaklikDouble);
                // İniş Hızı grafiğine veriyi ekleyin
                chart3.Series["InisHizi"].Points.Add(inisHiziDouble);
                // Nem grafiğine veriyi ekleyin
                chart5.Series["Nem"].Points.Add(nemDouble);
                // Basınç grafiğine veriyi ekleyin
                chart4.Series["Basinc"].Points.Add(basincDouble);
                // Irtifa grafiğine veriyi ekleyin
                chart6.Series["Irtifa"].Points.Add(irtifaDouble);
                // Sıcaklık grafiğine veriyi ekleyin
                chart7.Series["Nem"].Points.Add(sicaklikDouble);
                // Nem grafiğine veriyi ekleyin
                chart8.Series["Sicaklik"].Points.Add(nemDouble);
                // Basınç grafiğine veriyi ekleyin
                chart9.Series["Basinc"].Points.Add(basincDouble);
            }
            if (double.TryParse(rowData[9], out double intubePressure1Double))
            {
                // barITP1 circular barına veriyi ekleyin
                barITP1.Value = (int)intubePressure1Double;

                // Orta sayıyı güncelle (örneğin, bir Label kullanarak)
                barITP1.Text = intubePressure1Double.ToString(); // İntubePressure1 değerini ortasına yazdık
            }

            if (double.TryParse(rowData[10], out double intubePressure2Double))
            {
                // barITP2 circular barına veriyi ekleyin
                barITP2.Value = (int)intubePressure2Double;

                // Orta sayıyı güncelle (örneğin, bir Label kullanarak)
                barITP2.Text = intubePressure2Double.ToString(); // İntubePressure2 değerini ortasına yazdık
            }

            barAvoyonik.Value += 1;
            barAvoyonik.Text = barAvoyonik.Value.ToString();
            if (barAvoyonik.Value == 100)
            {
                timer1.Enabled = false;
            }
            currentRowIndex++;

        }

    }
}





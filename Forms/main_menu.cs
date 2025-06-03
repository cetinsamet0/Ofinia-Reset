using ofinia_reset.Stylers;
using System;
using System.Diagnostics;
using System.Runtime.ConstrainedExecution;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Text;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Window;



namespace ofinia_reset
{
    public partial class main_menu : Form
    {
        // Win32 API fonksiyonlar�n� �a��rmak i�in:
        [DllImport("user32.dll")]
        public static extern bool ReleaseCapture();

        [DllImport("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);

        // S�r�kleme i�in sabitler
        private const int WM_NCLBUTTONDOWN = 0xA1;
        private const int HTCAPTION = 0x2;
        public main_menu()
        {
            InitializeComponent();
            this.MouseDown += Form_MouseDown;

        }

        private void main_menu_Load(object sender, EventArgs e)
        {

            pictureBox2.Tag = "no-style";
            if (!IsRunningAsAdministrator())
            {
                MessageBox.Show("L�tfen program� y�netici olarak �al��t�r�n�z!!!", "Yetki Sorunu", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                Application.Exit();
            }

            PictureStyler.ApplyAllPictureBoxes(this);
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            Application.Exit();


        }



        public static bool IsRunningAsAdministrator()
        {
            using (WindowsIdentity identity = WindowsIdentity.GetCurrent())
            {
                WindowsPrincipal principal = new WindowsPrincipal(identity);
                return principal.IsInRole(WindowsBuiltInRole.Administrator);
            }
        }
        private void Form_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ReleaseCapture();
                SendMessage(this.Handle, WM_NCLBUTTONDOWN, HTCAPTION, 0);
            }
        }

        private void pictureBox4_Click(object sender, EventArgs e)
        {
            wordreset();
        }
        private void pictureBox5_Click(object sender, EventArgs e)
        {
            excelreset();
        }
        private void pictureBox6_Click(object sender, EventArgs e)
        {
            powerpointreset();
        }
        private void pictureBox7_Click(object sender, EventArgs e)
        {
            outlookreset();
        }


        public static void wordreset()
        {
            string wordScript = @"
                foreach ($v in '16.0','15.0','14.0','12.0') {
                    $key = 'HKCU:\Software\Microsoft\Office\' + $v + '\Word'
                    if (Test-Path $key) {
                        Remove-Item $key -Recurse -Force
                    }
                }
                $normalTemplate = $env:APPDATA + '\Microsoft\Templates\Normal.dotm'
                if (Test-Path $normalTemplate) {
                    Remove-Item $normalTemplate -Force
                }
                $recoveryPath = $env:APPDATA + '\Microsoft\Word'
                if (Test-Path $recoveryPath) {
                    Remove-Item $recoveryPath\* -Recurse -Force
                }
            ";
            DialogResult answer = MessageBox.Show("Microsoft Word uygulamas�n�n t�m ayarlar�n� s�f�rlamak istedi�inizden emin misiniz?", "Uyar�", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (answer == DialogResult.Yes)
            {
                try
                {
                    ProcessStartInfo psi = new ProcessStartInfo
                    {
                        FileName = "powershell.exe",
                        Arguments = $"-NoProfile -ExecutionPolicy Bypass -Command \"{wordScript}\"",
                        Verb = "runas", // Y�netici olarak �al��t�r
                        RedirectStandardOutput = true,
                        RedirectStandardError = true,
                        UseShellExecute = false,
                        CreateNoWindow = true,
                        StandardOutputEncoding = Encoding.UTF8
                    };
                    using (Process process = new Process())
                    {
                        process.StartInfo = psi;
                        process.Start();
                        process.WaitForExit();

                        if (process.ExitCode == 0)
                        {
                            MessageBox.Show("Word ayarlar� ba�ar�yla s�f�rland�!", "Tamamland�", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        else
                        {
                            string error = process.StandardError.ReadToEnd();
                            MessageBox.Show("Hata olu�tu:\n" + error, "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }


                        Process.Start(psi);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Komut �al��t�r�l�rken hata olu�tu:\n" + ex.Message);
                }

            }
            else
            {
                MessageBox.Show("S�f�rlama i�leminden vazge�ildi!", "�ptal Edildi");

            }


        }

        public static void powerpointreset()
        {
            string pptScript = @"
        foreach ($v in '16.0','15.0','14.0','12.0') {
            $key = 'HKCU:\Software\Microsoft\Office\' + $v + '\PowerPoint'
            if (Test-Path $key) {
                Remove-Item $key -Recurse -Force
            }
        }
        $templatePath = $env:APPDATA + '\Microsoft\Templates'
        $pptTemplates = Get-ChildItem -Path $templatePath -Filter '*.potx' -ErrorAction SilentlyContinue
        foreach ($file in $pptTemplates) {
            Remove-Item $file.FullName -Force
        }

        $recoveryPath = $env:APPDATA + '\Microsoft\PowerPoint'
        if (Test-Path $recoveryPath) {
            Remove-Item $recoveryPath\* -Recurse -Force
        }
    ";

            DialogResult answer = MessageBox.Show("Microsoft PowerPoint uygulamas�n�n t�m ayarlar�n� s�f�rlamak istedi�inizden emin misiniz?", "Uyar�", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (answer == DialogResult.Yes)
            {
                try
                {
                    ProcessStartInfo psi = new ProcessStartInfo
                    {
                        FileName = "powershell.exe",
                        Arguments = $"-NoProfile -ExecutionPolicy Bypass -Command \"{pptScript}\"",
                        Verb = "runas", // Y�netici olarak �al��t�r
                        RedirectStandardOutput = true,
                        RedirectStandardError = true,
                        UseShellExecute = false,
                        CreateNoWindow = true,
                        StandardOutputEncoding = Encoding.UTF8
                    };

                    using (Process process = new Process())
                    {
                        process.StartInfo = psi;
                        process.Start();
                        process.WaitForExit();

                        if (process.ExitCode == 0)
                        {
                            MessageBox.Show("PowerPoint ayarlar� ba�ar�yla s�f�rland�!", "Tamamland�", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        else
                        {
                            string error = process.StandardError.ReadToEnd();
                            MessageBox.Show("Hata olu�tu:\n" + error, "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Komut �al��t�r�l�rken hata olu�tu:\n" + ex.Message);
                }
            }
            else
            {
                MessageBox.Show("S�f�rlama i�leminden vazge�ildi!", "�ptal Edildi");
            }
        }

        public static void excelreset()
        {
            
        }

        public static void outlookreset()
        {
            string outlookScript = @"
        foreach ($v in '16.0','15.0','14.0','12.0') {
            $key = 'HKCU:\Software\Microsoft\Office\' + $v + '\Outlook'
            if (Test-Path $key) {
                Remove-Item $key -Recurse -Force
            }
        }

        # Profil ayarlar�
        $profileKey = 'HKCU:\Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles'
        if (Test-Path $profileKey) {
            Remove-Item $profileKey -Recurse -Force
        }

        # RoamCache (�nbellek verileri)
        $roamCache = $env:APPDATA + '\Microsoft\Outlook\RoamCache'
        if (Test-Path $roamCache) {
            Remove-Item $roamCache\* -Recurse -Force
        }

        # Auto-Complete listesi
        $autocompletePath = $env:APPDATA + '\Microsoft\Outlook'
        $autocompleteFiles = Get-ChildItem $autocompletePath -Filter '*.dat' -ErrorAction SilentlyContinue
        foreach ($file in $autocompleteFiles) {
            Remove-Item $file.FullName -Force
        }
    ";

            DialogResult answer = MessageBox.Show("Microsoft Outlook uygulamas�n�n t�m ayarlar�n� s�f�rlamak istedi�inizden emin misiniz?", "Uyar�", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (answer == DialogResult.Yes)
            {
                try
                {
                    ProcessStartInfo psi = new ProcessStartInfo
                    {
                        FileName = "powershell.exe",
                        Arguments = $"-NoProfile -ExecutionPolicy Bypass -Command \"{outlookScript}\"",
                        Verb = "runas", // Y�netici olarak �al��t�r
                        RedirectStandardOutput = true,
                        RedirectStandardError = true,
                        UseShellExecute = false,
                        CreateNoWindow = true,
                        StandardOutputEncoding = Encoding.UTF8
                    };

                    using (Process process = new Process())
                    {
                        process.StartInfo = psi;
                        process.Start();
                        process.WaitForExit();

                        if (process.ExitCode == 0)
                        {
                            MessageBox.Show("Outlook ayarlar� ba�ar�yla s�f�rland�!", "Tamamland�", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        else
                        {
                            string error = process.StandardError.ReadToEnd();
                            MessageBox.Show("Hata olu�tu:\n" + error, "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Komut �al��t�r�l�rken hata olu�tu:\n" + ex.Message);
                }
            }
            else
            {
                MessageBox.Show("S�f�rlama i�leminden vazge�ildi!", "�ptal Edildi");
            }
        }

        
    }
}

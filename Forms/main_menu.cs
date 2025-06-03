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
        // Win32 API fonksiyonlarýný çaðýrmak için:
        [DllImport("user32.dll")]
        public static extern bool ReleaseCapture();

        [DllImport("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);

        // Sürükleme için sabitler
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
                MessageBox.Show("Lütfen programý yönetici olarak çalýþtýrýnýz!!!", "Yetki Sorunu", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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
            DialogResult answer = MessageBox.Show("Microsoft Word uygulamasýnýn tüm ayarlarýný sýfýrlamak istediðinizden emin misiniz?", "Uyarý", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (answer == DialogResult.Yes)
            {
                try
                {
                    ProcessStartInfo psi = new ProcessStartInfo
                    {
                        FileName = "powershell.exe",
                        Arguments = $"-NoProfile -ExecutionPolicy Bypass -Command \"{wordScript}\"",
                        Verb = "runas", // Yönetici olarak çalýþtýr
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
                            MessageBox.Show("Word ayarlarý baþarýyla sýfýrlandý!", "Tamamlandý", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        else
                        {
                            string error = process.StandardError.ReadToEnd();
                            MessageBox.Show("Hata oluþtu:\n" + error, "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }


                        Process.Start(psi);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Komut çalýþtýrýlýrken hata oluþtu:\n" + ex.Message);
                }

            }
            else
            {
                MessageBox.Show("Sýfýrlama iþleminden vazgeçildi!", "Ýptal Edildi");

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

            DialogResult answer = MessageBox.Show("Microsoft PowerPoint uygulamasýnýn tüm ayarlarýný sýfýrlamak istediðinizden emin misiniz?", "Uyarý", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (answer == DialogResult.Yes)
            {
                try
                {
                    ProcessStartInfo psi = new ProcessStartInfo
                    {
                        FileName = "powershell.exe",
                        Arguments = $"-NoProfile -ExecutionPolicy Bypass -Command \"{pptScript}\"",
                        Verb = "runas", // Yönetici olarak çalýþtýr
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
                            MessageBox.Show("PowerPoint ayarlarý baþarýyla sýfýrlandý!", "Tamamlandý", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        else
                        {
                            string error = process.StandardError.ReadToEnd();
                            MessageBox.Show("Hata oluþtu:\n" + error, "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Komut çalýþtýrýlýrken hata oluþtu:\n" + ex.Message);
                }
            }
            else
            {
                MessageBox.Show("Sýfýrlama iþleminden vazgeçildi!", "Ýptal Edildi");
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

        # Profil ayarlarý
        $profileKey = 'HKCU:\Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles'
        if (Test-Path $profileKey) {
            Remove-Item $profileKey -Recurse -Force
        }

        # RoamCache (önbellek verileri)
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

            DialogResult answer = MessageBox.Show("Microsoft Outlook uygulamasýnýn tüm ayarlarýný sýfýrlamak istediðinizden emin misiniz?", "Uyarý", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (answer == DialogResult.Yes)
            {
                try
                {
                    ProcessStartInfo psi = new ProcessStartInfo
                    {
                        FileName = "powershell.exe",
                        Arguments = $"-NoProfile -ExecutionPolicy Bypass -Command \"{outlookScript}\"",
                        Verb = "runas", // Yönetici olarak çalýþtýr
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
                            MessageBox.Show("Outlook ayarlarý baþarýyla sýfýrlandý!", "Tamamlandý", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        else
                        {
                            string error = process.StandardError.ReadToEnd();
                            MessageBox.Show("Hata oluþtu:\n" + error, "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Komut çalýþtýrýlýrken hata oluþtu:\n" + ex.Message);
                }
            }
            else
            {
                MessageBox.Show("Sýfýrlama iþleminden vazgeçildi!", "Ýptal Edildi");
            }
        }

        
    }
}

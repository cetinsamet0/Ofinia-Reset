using System;
using System.Security.Principal;


namespace ofinia_reset
{
    public partial class main_menu : Form
    {
        public main_menu()
        {
            InitializeComponent();
        }
        
        private void main_menu_Load(object sender, EventArgs e)
        {
            //if (!IsRunningAsAdministrator())
            //{
            //    MessageBox.Show("L�tfen program� y�netici olarak �al��t�r�n�z!!!","Yetki Sorunu",MessageBoxButtons.OK,MessageBoxIcon.Warning);
            //    Application.Exit();
            //}
           

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
    }
}

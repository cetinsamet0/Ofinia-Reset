using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Window;



namespace ofinia_reset.Stylers
{
    internal class PictureStyler
    {


        public static void ApplyStyle(PictureBox pictureBox)
        {      
            if (pictureBox.Tag != null && pictureBox.Tag.ToString() == "no-style") 
                return;
            pictureBox.Cursor = Cursors.Hand;
            pictureBox.MouseEnter -= OnMouseEnter;
            pictureBox.MouseLeave -= OnMouseLeave;
            pictureBox.MouseEnter += OnMouseEnter;
            pictureBox.MouseLeave += OnMouseLeave;
            pictureBox.MouseDown += MouseDown;
            

        }
        public static void ApplyAllPictureBoxes(Control parent)
        {
            foreach (Control control in parent.Controls)
            {

                if (control is PictureBox pb) { 

                 ApplyStyle(pb);

                }

            }

        }
        private static void OnMouseEnter(object sender,EventArgs e) { 
        if (sender is PictureBox pb)
            {
                pb.BorderStyle = BorderStyle.FixedSingle;
            }
            
        }
        private static void OnMouseLeave(object sender,EventArgs e) {
            if (sender is PictureBox pb)
            {
                pb.BorderStyle = BorderStyle.None;
            }
        }
        private static void MouseDown(object sender,EventArgs e)
        {
            if (sender is PictureBox pb)
            {
                pb.BorderStyle = BorderStyle.Fixed3D;
            }
        }
        

       
    }
}

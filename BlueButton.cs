using System.Drawing;
using System.Windows.Forms;

namespace MartinAppGUI
{
    public partial class BlueButton : Button
    {
        public BlueButton()
        {
            InitializeComponent();

            ForeColor = Color.White;
            CurrentBackColor = Color.FromArgb(30, 50, 94);
        }

        private Color CurrentBackColor;

        private Color onHoverBackColor = Color.FromArgb(30, 84, 161);

        public Color OnHoverBackColor
        {
            get { return onHoverBackColor; }
            set { onHoverBackColor = value; Invalidate(); }
        }
    }
}

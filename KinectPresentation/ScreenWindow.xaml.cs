using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace KinectPresentation
{
    /// <summary>
    /// ScreenWindow.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class ScreenWindow : Window
    {
        /// <summary>
        /// Drawing group for laser point
        /// </summary>
        public DrawingGroup drawingGroup;

        /// <summary>
        /// Drawing image that we will display
        /// </summary>
        private DrawingImage imageSource;

        /// <summary>
        /// Red pen to use lase
        /// </summary>
        private System.Windows.Media.Pen redPen = new System.Windows.Media.Pen(System.Windows.Media.Brushes.Red, 1);

        /// <summary>
        /// Left/Right laser enumeration
        /// </summary>
        public enum LaserType
        {
            LASER_LEFT = 1,
            LASER_RIGHT = 2
        };

        public ScreenWindow()
        {
            InitializeComponent();

            // Resize to screen resolution
            System.Drawing.Size WorkingArea = Screen.PrimaryScreen.WorkingArea.Size;

            this.Top = 0;
            this.Left = 0;
            imgBoard.Width = this.Width = WorkingArea.Width;
            imgBoard.Height = this.Height = WorkingArea.Height;
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            // Create the drawing group we'll use for drawing
            this.drawingGroup = new DrawingGroup();

            // Create an image source that we can use in our image control
            this.imageSource = new DrawingImage(this.drawingGroup);

            // Display the drawing using our image control
            imgBoard.Source = this.imageSource;
        }

        public void DrawLaserPoint(LaserType laserType, DrawingContext dc, System.Windows.Point point)
        {
            dc.DrawRectangle(System.Windows.Media.Brushes.Transparent, null, new Rect(0, 0, imgBoard.Width, imgBoard.Height));
            if (point.X == 0 && point.Y == 0)
            {
                // Erase(Fake draw... zzz)
                dc.DrawEllipse(System.Windows.Media.Brushes.Transparent,
                                new System.Windows.Media.Pen(System.Windows.Media.Brushes.Transparent, 1),
                                new System.Windows.Point(0, 0), 0, 0);
            }
            else
            {
                // Draw
                dc.DrawEllipse(System.Windows.Media.Brushes.Red, redPen, new System.Windows.Point(point.X, point.Y), 2, 2);
            }
        }
    }
}

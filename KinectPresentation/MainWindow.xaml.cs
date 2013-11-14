using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

using Microsoft.Kinect;
using Microsoft.Samples.Kinect.SwipeGestureRecognizer;

namespace KinectPresentation
{
    /// <summary>
    /// MainWindow.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class MainWindow : Window
    {
        /// <summary>
        /// Width of output drawing
        /// </summary>
        private const float RenderWidth = 640.0f;

        /// <summary>
        /// Height of our output drawing
        /// </summary>
        private const float RenderHeight = 480.0f;

        /// <summary>
        /// Thickness of drawn joint lines
        /// </summary>
        private const double JointThickness = 1;

        /// <summary>
        /// Thickness of body center ellipse
        /// </summary>
        private const double BodyCenterThickness = 1;

        /// <summary>
        /// Thickness of clip edge rectangles
        /// </summary>
        private const double ClipBoundsThickness = 10;

        /// <summary>
        /// Brush used to draw skeleton center point
        /// </summary>
        private readonly Brush centerPointBrush = Brushes.Blue;

        /// <summary>
        /// Brush used for drawing joints that are currently inferred
        /// </summary>        
        private readonly Brush inferredJointBrush = Brushes.Yellow;

        /// <summary>
        /// Pen used for drawing bones that are currently tracked
        /// </summary>
        private readonly Pen trackedBonePen = new Pen(Brushes.Green, 6);

        /// <summary>
        /// Pen used for drawing bones that are currently inferred
        /// </summary>        
        private readonly Pen inferredBonePen = new Pen(Brushes.Gray, 1);

        /// <summary>
        /// Active Kinect sensor
        /// </summary>
        private KinectSensor sensor;

        /// <summary>
        /// Drawing group for skeleton rendering output
        /// </summary>
        private DrawingGroup drawingGroup;

        /// <summary>
        /// Drawing image that we will display
        /// </summary>
        private DrawingImage imageSource;

        /// <summary>
        /// Window used for drawing pen or laser
        /// </summary>
        private ScreenWindow wndScreen;

        /// <summary>
        /// Screen window handle
        /// </summary>
        private IntPtr hwndScreen = (System.IntPtr)null;

        /// <summary>
        /// Presentation window handle
        /// </summary>
        private IntPtr hwndPresentation = (System.IntPtr)null;

        /// <summary>
        /// Presentation window informations
        /// </summary>
        private IList windowInfo = new ArrayList();

        /// <summary>
        /// Presentation window informations' index enum
        /// </summary>
        enum wndInfo
        {
            Application = 0,
            WindowCaption = 1,
            ClassName = 2
        };

        /// <summary>
        /// Retrieves a handle to the top-level window whose class name and window name match the specified strings.
        /// </summary>
        [DllImport("user32.dll")]
        public static extern bool EnumChildWindows(IntPtr hwndParent, EnumWindowsProc lpEnumFunc, ref IList lParam);

        /// <summary>
        /// Retrieves the name of the class to which the specified window belongs.
        /// </summary>
        [DllImport("User32.dll", EntryPoint = "GetClassNameW", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern IntPtr GetClassName(IntPtr hWnd, [MarshalAs(UnmanagedType.LPWStr)] StringBuilder lpClassName, int nMaxCount);

        /// <summary>
        /// Copies the text of the specified window's title bar (if it has one) into a buffer.
        /// </summary>
        [DllImport("user32.dll")]
        public static extern int GetWindowText(int hwnd, StringBuilder buf, int nMaxCount);

        /// <summary>
        /// The EnumChildWindows' callback function.
        /// </summary>
        public delegate int EnumWindowsProc(IntPtr hwnd, ref IList lParam);

        /// <summary>
        /// Places (posts) a message in the message queue associated with the thread that created the specified window and returns without waiting for the thread to process the message.
        /// </summary>
        [DllImport("user32.dll")]
        private static extern int PostMessage(IntPtr hWnd, uint Msg, uint wParam, uint lParam);

        /// <summary>
        /// Brings the thread that created the specified window into the foreground and activates the window.
        /// </summary>
        [DllImport("User32.dll")]
        private static extern int SetForegroundWindow(IntPtr hwnd);

        /// <summary>
        /// Keydown message constant.
        /// </summary>
        private const int WM_KEYDOWN = 0x100;

        /// <summary>
        /// Keyup message constant.
        /// </summary>
        private const int WM_KEYUP = 0x101;

        /// <summary>
        /// The virture key constant of left direction key.
        /// </summary>
        private const int VK_LEFT = 0x25;

        /// <summary>
        /// The virture key constant of right direction key.
        /// </summary>
        private const int VK_RIGHT = 0x27;

        /// <summary>
        /// Recognizer for detecting swipe
        /// </summary>
        private Recognizer activeRecognizer;

        /// <summary>
        /// Array to receive skeletons from sensor, resize when needed.
        /// </summary>
        private Skeleton[] skeletons = new Skeleton[0];

        /// <summary>
        /// Initializes a new instance of the MainWindow class.
        /// </summary>
        public MainWindow()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Execute startup tasks
        /// </summary>
        /// <param name="sender">object sending the event</param>
        /// <param name="e">event arguments</param>
        private void WindowLoaded(object sender, RoutedEventArgs e)
        {
            // Create the gesture recognizer.
            this.activeRecognizer = this.CreateRecognizer();

            // Create the drawing group we'll use for drawing
            this.drawingGroup = new DrawingGroup();

            // Create an image source that we can use in our image control
            this.imageSource = new DrawingImage(this.drawingGroup);

            // Display the drawing using our image control
            Image.Source = this.imageSource;

            // Look through all sensors and start the first connected one.
            // This requires that a Kinect is connected at the time of app startup.
            // To make your app robust against plug/unplug, 
            // it is recommended to use KinectSensorChooser provided in Microsoft.Kinect.Toolkit
            foreach (var potentialSensor in KinectSensor.KinectSensors)
            {
                if (potentialSensor.Status == KinectStatus.Connected)
                {
                    this.sensor = potentialSensor;
                    break;
                }
            }

            if (null != this.sensor)
            {
                // Turn on the skeleton stream to receive skeleton frames
                this.sensor.SkeletonStream.Enable();

                // Add an event handler to be called whenever there is new color frame data
                this.sensor.SkeletonFrameReady += this.SensorSkeletonFrameReady;

                // Start the sensor!
                try
                {
                    this.sensor.Start();
                }
                catch (IOException)
                {
                    this.sensor = null;
                }
            }

            if (null == this.sensor)
            {
                this.txtStatus.Text = "No kinect has detected.";
            }

            // Show screen window to draw pen or laser
            wndScreen = new ScreenWindow();
            wndScreen.Show();

            InitializeWindowInfo();

            hwndPresentation = FindPresentationHandle((IntPtr)null, this.windowInfo);
            if (hwndPresentation == (IntPtr)null)
            {
                this.txtStatus.Text = "No presentation application has detected.";
            }

            // Find windows handle to send message
            hwndScreen = new WindowInteropHelper(this.wndScreen).Handle;
        }

        /// <summary>
        /// Create recognizer to detect hand swaping motion
        /// </summary>
        private Recognizer CreateRecognizer()
        {
            // Instantiate a recognizer.
            var recognizer = new Recognizer();

            // Wire-up swipe left to manually reverse picture.
            recognizer.SwipeLeftDetected += (s, e) =>
            {
                // Process prev
                this.txtStatus.Text = string.Format("Left arm swiped!!!");
                SetForegroundWindow(this.hwndPresentation);
                PostMessage(this.hwndPresentation, WM_KEYDOWN, VK_LEFT, 0x14B0001);
                PostMessage(this.hwndPresentation, WM_KEYUP, VK_LEFT, 0xC14B0001);
                //SetForegroundWindow(this.hwndScreen);
            };

            // Wire-up swipe right to manually advance picture.
            recognizer.SwipeRightDetected += (s, e) =>
            {
                // Process next
                this.txtStatus.Text = string.Format("Right arm swiped!!!");
                SetForegroundWindow(this.hwndPresentation);
                PostMessage(this.hwndPresentation, WM_KEYDOWN, VK_RIGHT, 0x14D0001);
                PostMessage(this.hwndPresentation, WM_KEYUP, VK_RIGHT, 0xC14D0001);
                //SetForegroundWindow(this.hwndScreen);
            };

            return recognizer;
        }

        /// <summary>
        /// Set the presentation's window information.
        /// </summary>
        private void InitializeWindowInfo()
        {
            EnumerateWindow enumWindow;
            EnumerateWindow enumChildWindow;

            // HanShow 2010
            enumWindow = new EnumerateWindow(0, "슬라이드 쇼", "CHslShowView");
            windowInfo.Add(enumWindow);

            // Microsoft PowerPoint 2007
            enumWindow = new EnumerateWindow(0, "", "screenClass");
            enumChildWindow = enumWindow;
            enumChildWindow = enumChildWindow.AddChild(0, "슬라이드 쇼", "PaneClassDC");
            windowInfo.Add(enumWindow);

            // SlideShare
            enumWindow = new EnumerateWindow(0, "Upload & Share PowerPoint presentations and documents - Chrome", "Chrome_WidgetWin_1");
            enumChildWindow = enumWindow;
            enumChildWindow = enumChildWindow.AddChild(0, "Upload & Share PowerPoint presentations and documents", "Chrome_WidgetWin_0");
            enumChildWindow = enumChildWindow.AddChild(0, "", "Chrome_RenderWidgetHostHWND");
            windowInfo.Add(enumWindow);
        }

        /// <summary>
        /// Execute shutdown tasks
        /// </summary>
        /// <param name="sender">object sending the event</param>
        /// <param name="e">event arguments</param>
        private void WindowClosing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            this.wndScreen.Close();

            if (null != this.sensor)
            {
                this.sensor.Stop();
            }
        }

        /// <summary>
        /// Draws indicators to show which edges are clipping skeleton data
        /// </summary>
        /// <param name="skeleton">skeleton to draw clipping information for</param>
        /// <param name="drawingContext">drawing context to draw to</param>
        private static void RenderClippedEdges(Skeleton skeleton, DrawingContext drawingContext)
        {
            if (skeleton.ClippedEdges.HasFlag(FrameEdges.Bottom))
            {
                drawingContext.DrawRectangle(
                    Brushes.Red,
                    null,
                    new Rect(0, RenderHeight - ClipBoundsThickness, RenderWidth, ClipBoundsThickness));
            }

            if (skeleton.ClippedEdges.HasFlag(FrameEdges.Top))
            {
                drawingContext.DrawRectangle(
                    Brushes.Red,
                    null,
                    new Rect(0, 0, RenderWidth, ClipBoundsThickness));
            }

            if (skeleton.ClippedEdges.HasFlag(FrameEdges.Left))
            {
                drawingContext.DrawRectangle(
                    Brushes.Red,
                    null,
                    new Rect(0, 0, ClipBoundsThickness, RenderHeight));
            }

            if (skeleton.ClippedEdges.HasFlag(FrameEdges.Right))
            {
                drawingContext.DrawRectangle(
                    Brushes.Red,
                    null,
                    new Rect(RenderWidth - ClipBoundsThickness, 0, ClipBoundsThickness, RenderHeight));
            }
        }

        /// <summary>
        /// Find presentation window.
        /// </summary>
        /// <param name="hwndParent">Parent window's handle</param>
        /// <param name="windowInfo">Comparing data</param>
        /// <returns></returns>
        private IntPtr FindPresentationHandle(IntPtr hwndParent, IList windowInfo)
        {
            IntPtr hwndPresentation = (IntPtr)null;
            IList enumWindows = new ArrayList();

            EnumChildWindows(hwndParent, new EnumWindowsProc(WindowEnum), ref enumWindows);

            foreach (EnumerateWindow enumWind in enumWindows)
            {
                if (enumWind.handle != (IntPtr)null)
                {
                    foreach (EnumerateWindow enumWindInfo in windowInfo)
                    {
                        if (enumWind.caption.ToString() == enumWindInfo.caption.ToString() ||
                            enumWind.className.ToString() == enumWindInfo.className.ToString())
                        {
                            if (enumWindInfo.hasChild() && 0 < enumWindInfo.childWindows.Count)
                            {
                                hwndPresentation = FindPresentationHandle(enumWind.handle, enumWindInfo.childWindows);
                            }
                            else
                            {
                                hwndPresentation = enumWind.handle;
                            }

                            if (hwndPresentation != (IntPtr)null)
                            {
                                return hwndPresentation;
                            }
                        }
                    }
                }
            }

            return hwndPresentation;
        }

        /// <summary>
        /// The EnumChildWindows' callback function.
        /// </summary>
        /// <param name="hWnd">Window handle</param>
        /// <param name="lParam">Additional data(In this, IList for saving data)</param>
        /// <returns></returns>
        private static int WindowEnum(IntPtr hWnd, ref IList lParam)
        {
            EnumerateWindow enumWindow = new EnumerateWindow();
            StringBuilder className = new StringBuilder(255);
            StringBuilder caption = new StringBuilder(255);

            GetClassName(hWnd, className, className.Capacity);
            GetWindowText(hWnd.ToInt32(), caption, caption.Capacity);

            enumWindow.handle = hWnd;
            enumWindow.className = className;
            enumWindow.caption = caption;
            lParam.Add(enumWindow);

            return 1;
        }

        /// <summary>
        /// Event handler for Kinect sensor's SkeletonFrameReady event
        /// </summary>
        /// <param name="sender">object sending the event</param>
        /// <param name="e">event arguments</param>
        private void SensorSkeletonFrameReady(object sender, SkeletonFrameReadyEventArgs e)
        {
            //Skeleton[] skeletons = new Skeleton[0];
            DepthImagePoint wristDepthPoint, elbowDepthPoint, boardDepthPoint;

            using (SkeletonFrame skeletonFrame = e.OpenSkeletonFrame())
            {
                if (skeletonFrame != null)
                {
                    this.skeletons = new Skeleton[skeletonFrame.SkeletonArrayLength];
                    skeletonFrame.CopySkeletonDataTo(this.skeletons);

                    // Pass skeletons to recognizer.
                    this.activeRecognizer.Recognize(sender, skeletonFrame, this.skeletons);
                }
            }

            using (DrawingContext dc = this.drawingGroup.Open())
            {
                // Draw a transparent background to set the render size
                dc.DrawRectangle(Brushes.Black, null, new Rect(new Point(0, 0), new Point(RenderWidth, RenderHeight)));

                if (this.skeletons.Length != 0)
                {
                    foreach (Skeleton skel in this.skeletons)
                    {
                        //RenderClippedEdges(skel, dc);

                        if (skel.TrackingState == SkeletonTrackingState.Tracked)
                        {
                            this.DrawBones(skel, dc);

                            using (DrawingContext laserDC = wndScreen.drawingGroup.Open())
                            {

                                // Left Arm
                                wristDepthPoint = SkeletonPointToScreenWithDepth(skel.Joints[JointType.WristLeft].Position);
                                this.txtLeftWrist.Text = string.Format("({0}, {1}, {2})", wristDepthPoint.X, wristDepthPoint.Y, wristDepthPoint.Depth);
                                elbowDepthPoint = SkeletonPointToScreenWithDepth(skel.Joints[JointType.HandLeft].Position);
                                this.txtLeftElbow.Text = string.Format("({0}, {1}, {2})", elbowDepthPoint.X, elbowDepthPoint.Y, elbowDepthPoint.Depth);
                                boardDepthPoint = GetPointOnBoard(0, wristDepthPoint, elbowDepthPoint);
                                this.txtLeftPoint.Text = string.Format("({0}, {1}, {2})", boardDepthPoint.X, boardDepthPoint.Y, boardDepthPoint.Depth);

                                if ((wristDepthPoint.Depth > elbowDepthPoint.Depth) &&
                                    (0 < boardDepthPoint.X && boardDepthPoint.X < RenderWidth) &&
                                    (0 < boardDepthPoint.Y && boardDepthPoint.Y < RenderHeight))
                                {
                                    this.txtStatus.Text = string.Format("Draw left laser point on ({0}, {1}, {2})",
                                                                        boardDepthPoint.X, boardDepthPoint.Y, boardDepthPoint.Depth);
                                    wndScreen.DrawLaserPoint(ScreenWindow.LaserType.LASER_LEFT, laserDC,
                                                            new Point(boardDepthPoint.X * (wndScreen.Width / RenderWidth), boardDepthPoint.Y * (wndScreen.Height / RenderHeight)));
                                }
                                else
                                {
                                    wndScreen.DrawLaserPoint(ScreenWindow.LaserType.LASER_LEFT, laserDC, new Point(0, 0));
                                }

                                // Right Arm
                                wristDepthPoint = SkeletonPointToScreenWithDepth(skel.Joints[JointType.WristRight].Position);
                                this.txtRightWrist.Text = string.Format("({0}, {1}, {2})", wristDepthPoint.X, wristDepthPoint.Y, wristDepthPoint.Depth);
                                elbowDepthPoint = SkeletonPointToScreenWithDepth(skel.Joints[JointType.HandRight].Position);
                                this.txtRightElbow.Text = string.Format("({0}, {1}, {2})", elbowDepthPoint.X, elbowDepthPoint.Y, elbowDepthPoint.Depth);
                                boardDepthPoint = GetPointOnBoard(0, wristDepthPoint, elbowDepthPoint);
                                this.txtRightPoint.Text = string.Format("({0}, {1}, {2})", boardDepthPoint.X, boardDepthPoint.Y, boardDepthPoint.Depth);

                                if ((wristDepthPoint.Depth > elbowDepthPoint.Depth) &&
                                    (0 < boardDepthPoint.X && boardDepthPoint.X < RenderWidth) &&
                                    (0 < boardDepthPoint.Y && boardDepthPoint.Y < RenderHeight))
                                {
                                    this.txtStatus.Text = string.Format("Draw right laser point on ({0}, {1}, {2})",
                                                                        boardDepthPoint.X, boardDepthPoint.Y, boardDepthPoint.Depth);
                                    wndScreen.DrawLaserPoint(ScreenWindow.LaserType.LASER_RIGHT, laserDC,
                                                            new Point(boardDepthPoint.X * (wndScreen.Width / RenderWidth), boardDepthPoint.Y * (wndScreen.Height / RenderHeight)));
                                }
                                else
                                {
                                    wndScreen.DrawLaserPoint(ScreenWindow.LaserType.LASER_RIGHT, laserDC, new Point(0, 0));
                                }
                            }
                        }
                    }
                }

                // prevent drawing outside of our render area
                this.drawingGroup.ClipGeometry = new RectangleGeometry(new Rect(0.0, 0.0, RenderWidth, RenderHeight));
            }
        }

        /// <summary>
        /// Get Point assumed on board using 3d linear equation
        /// </summary>
        /// <param name="assumed">the depth assumed</param>
        /// <param name="depthPoint0">depthPoint to start getting point on board from</param>
        /// <param name="depthPoint1">depthPoint to end getting point on board at</param>
        /// <returns>Getted depthPoint on board</returns>
        private DepthImagePoint GetPointOnBoard(int assumedDepth, DepthImagePoint depthPoint0, DepthImagePoint depthPoint1)
        {
            DepthImagePoint depthPoint = depthPoint0;
            int t;

            if (depthPoint1.Depth - depthPoint0.Depth != 0)
            {
                t = (assumedDepth - depthPoint0.Depth) / (depthPoint1.Depth - depthPoint0.Depth);

                depthPoint.X = t * (depthPoint1.X - depthPoint0.X) + depthPoint0.X;
                depthPoint.Y = t * (depthPoint1.Y - depthPoint0.Y) + depthPoint0.Y;
                depthPoint.Depth = assumedDepth;
            }
            else
            {
                depthPoint.X = 0;
                depthPoint.Y = 0;
                depthPoint.Depth = 0;
            }

            return depthPoint;
        }

        /// <summary>
        /// Maps a SkeletonPoint to line within our render space and converts to Point
        /// </summary>
        /// <param name="skelpoint">point to map</param>
        /// <returns>mapped point</returns>
        private Point SkeletonPointToScreen(SkeletonPoint skelpoint)
        {
            // Convert point to depth space.
            // We are not using depth directly, but we do want the points in our 640x480 output resolution.
            DepthImagePoint depthPoint = this.sensor.CoordinateMapper.MapSkeletonPointToDepthPoint(skelpoint, DepthImageFormat.Resolution640x480Fps30);
            return new Point(depthPoint.X, depthPoint.Y);
        }

        /// <summary>
        /// Maps a SkeletonPoint to line within our render space and converts to depthPoint
        /// </summary>
        /// <param name="skelpoint">depthPoint to map</param>
        /// <returns>mapped depthPoint</returns>
        private DepthImagePoint SkeletonPointToScreenWithDepth(SkeletonPoint skelpoint)
        {
            // Convert point to depth space.
            // we do want the points in our 640x480 output resolution.
            DepthImagePoint depthPoint = this.sensor.CoordinateMapper.MapSkeletonPointToDepthPoint(skelpoint, DepthImageFormat.Resolution640x480Fps30);
            return depthPoint;
        }

        /// <summary>
        /// Draws a skeleton's bones and joints
        /// </summary>
        /// <param name="skeleton">skeleton to draw</param>
        /// <param name="drawingContext">drawing context to draw to</param>
        private void DrawBones(Skeleton skeleton, DrawingContext drawingContext)
        {
            // Left Arm
            this.DrawBone(skeleton, drawingContext, JointType.ElbowLeft, JointType.WristLeft);
            //this.DrawBone(skeleton, drawingContext, JointType.WristLeft, JointType.HandLeft);

            // Right Arm
            this.DrawBone(skeleton, drawingContext, JointType.ElbowRight, JointType.WristRight);
            //this.DrawBone(skeleton, drawingContext, JointType.WristRight, JointType.HandRight);
        }

        /// <summary>
        /// Draws a bone line between two joints
        /// </summary>
        /// <param name="skeleton">skeleton to draw bones from</param>
        /// <param name="drawingContext">drawing context to draw to</param>
        /// <param name="jointType0">joint to start drawing from</param>
        /// <param name="jointType1">joint to end drawing at</param>
        private void DrawBone(Skeleton skeleton, DrawingContext drawingContext, JointType jointType0, JointType jointType1)
        {
            Joint joint0 = skeleton.Joints[jointType0];
            Joint joint1 = skeleton.Joints[jointType1];

            // If we can't find either of these joints, exit
            if (joint0.TrackingState == JointTrackingState.NotTracked ||
                joint1.TrackingState == JointTrackingState.NotTracked)
            {
                return;
            }

            // Don't draw if both points are inferred
            if (joint0.TrackingState == JointTrackingState.Inferred &&
                joint1.TrackingState == JointTrackingState.Inferred)
            {
                return;
            }

            // We assume all drawn bones are inferred unless BOTH joints are tracked
            Pen drawPen = this.inferredBonePen;
            if (joint0.TrackingState == JointTrackingState.Tracked && joint1.TrackingState == JointTrackingState.Tracked)
            {
                drawPen = this.trackedBonePen;
            }

            drawingContext.DrawLine(drawPen, this.SkeletonPointToScreen(joint0.Position), this.SkeletonPointToScreen(joint1.Position));
        }
    }
}

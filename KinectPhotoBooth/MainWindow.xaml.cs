using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using Microsoft.Kinect;
using System.IO;
using System.Collections.Generic;
using System.Windows.Media.Effects;
using AForge.Video;
using AForge.Video.DirectShow;
using System.Drawing;
using System.Windows.Media.Animation;

// Resolve ambiguities
using Brushes = System.Windows.Media.Brushes;
using FontFamily = System.Windows.Media.FontFamily;
using Point = System.Windows.Point;

namespace KinectPhotobooth
{
    public partial class MainWindow : Window
    {
        private KinectSensor kinectSensor;
        private WriteableBitmap colorBitmap;
        private byte[] colorPixels;
        private Skeleton[] skeletonData;
        private Button goButton;
        private TextBlock messageBlock;
        private TextBlock countdownBlock;
        private DateTime handOverButtonStartTime;
        private const double HandOverButtonThreshold = 1.5; // seconds
        private DispatcherTimer countdownTimer;
        private int photoCount = 0;
        private List<BitmapSource> capturedImages = new List<BitmapSource>();
        private bool isHandOverButton = false;

        private FilterInfoCollection videoDevices;
        private VideoCaptureDevice webcam;
        private System.Drawing.Bitmap currentWebcamFrame;

        public MainWindow()
        {
            InitializeComponent();
            InitializeKinect();
            InitializeWebcam();
            CreateUI();
            SetupCountdownTimer();

            KeyDown += MainWindow_KeyDown;
        }

        private void InitializeKinect()
        {
            kinectSensor = KinectSensor.KinectSensors[0];
            if (kinectSensor.Status == KinectStatus.Connected)
            {
                kinectSensor.ColorStream.Enable(ColorImageFormat.RgbResolution640x480Fps30);
                kinectSensor.SkeletonStream.Enable();

                colorPixels = new byte[kinectSensor.ColorStream.FramePixelDataLength];
                colorBitmap = new WriteableBitmap(
                    kinectSensor.ColorStream.FrameWidth,
                    kinectSensor.ColorStream.FrameHeight,
                    96.0, 96.0, PixelFormats.Bgr32, null);

                kinectImage.Source = colorBitmap;

                kinectSensor.ColorFrameReady += KinectSensor_ColorFrameReady;
                kinectSensor.SkeletonFrameReady += KinectSensor_SkeletonFrameReady;
                kinectSensor.Start();
            }
        }

        private void InitializeWebcam()
        {
            videoDevices = new FilterInfoCollection(FilterCategory.VideoInputDevice);
            if (videoDevices.Count > 0)
            {
                webcam = new VideoCaptureDevice(videoDevices[0].MonikerString);
                webcam.NewFrame += Webcam_NewFrame;
                webcam.Start();
            }
            else
            {
                MessageBox.Show("No webcam detected. Using Kinect camera for photos.");
            }
        }

        private void Webcam_NewFrame(object sender, NewFrameEventArgs eventArgs)
        {
            currentWebcamFrame = (System.Drawing.Bitmap)eventArgs.Frame.Clone();
        }

        private void CreateUI()
        {
            // Create GO! button
            goButton = new Button
            {
                Content = "Start",
                Width = 200,
                Height = 200,
                FontSize = 48,
                FontWeight = FontWeights.Bold,
                Style = (Style)FindResource("WeddingButton")
            };
            Canvas.SetLeft(goButton, (ActualWidth - goButton.Width) / 2);
            Canvas.SetTop(goButton, ActualHeight * 0.2); // Position at 20% from the top
            mainCanvas.Children.Add(goButton);

            // Create message block
            messageBlock = new TextBlock
            {
                FontSize = 36,
                Foreground = Brushes.White,
                TextAlignment = TextAlignment.Center,
                Width = 1000,
                TextWrapping = TextWrapping.Wrap,
                FontFamily = new FontFamily("Segoe Script"),
                Effect = new DropShadowEffect
                {
                    Color = Colors.Black,
                    Direction = 320,
                    ShadowDepth = 5,
                    Opacity = 0.5
                }
            };
            Canvas.SetLeft(messageBlock, (ActualWidth - messageBlock.Width) / 2);
            Canvas.SetTop(messageBlock, ActualHeight - 200);
            mainCanvas.Children.Add(messageBlock);

            // Create countdown block
            countdownBlock = new TextBlock
            {
                FontSize = 200,
                Foreground = Brushes.White,
                TextAlignment = TextAlignment.Center,
                Width = 300,
                FontWeight = FontWeights.Bold,
                Effect = new DropShadowEffect
                {
                    Color = Colors.Black,
                    Direction = 320,
                    ShadowDepth = 5,
                    Opacity = 0.5
                }
            };
            Canvas.SetLeft(countdownBlock, (ActualWidth - countdownBlock.Width) / 2);
            Canvas.SetTop(countdownBlock, (ActualHeight - countdownBlock.FontSize) / 2);
            mainCanvas.Children.Add(countdownBlock);

            ShowWelcomeMessage();

            // Handle window resize
            SizeChanged += MainWindow_SizeChanged;
        }

        private void MainWindow_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Escape)
            {
                Close();
            }
        }

        private void MainWindow_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            Canvas.SetLeft(goButton, (ActualWidth - goButton.Width) / 2);
            Canvas.SetTop(goButton, ActualHeight * 0.2);
            Canvas.SetLeft(messageBlock, (ActualWidth - messageBlock.Width) / 2);
            Canvas.SetTop(messageBlock, ActualHeight - 200);
            Canvas.SetLeft(countdownBlock, (ActualWidth - countdownBlock.Width) / 2);
            Canvas.SetTop(countdownBlock, (ActualHeight - countdownBlock.FontSize) / 2);
        }

        private void ShowWelcomeMessage()
        {
            messageBlock.Text = "Welcome to K+D's Wedding Photobooth!\nHold your hand over the Start button to begin!";
        }

        private void SetupCountdownTimer()
        {
            countdownTimer = new DispatcherTimer();
            countdownTimer.Interval = TimeSpan.FromSeconds(1);
            countdownTimer.Tick += CountdownTimer_Tick;
        }

        private void CountdownTimer_Tick(object sender, EventArgs e)
        {
            int currentCount = int.Parse(countdownBlock.Text);
            if (currentCount > 1)
            {
                countdownBlock.Text = (currentCount - 1).ToString();
            }
            else
            {
                countdownTimer.Stop();
                CapturePhoto();
            }
        }

        private void CapturePhoto()
        {
            BitmapSource capturedImage;

            if (webcam != null && currentWebcamFrame != null)
            {
                capturedImage = ConvertBitmapToBitmapSource(currentWebcamFrame);
            }
            else
            {
                capturedImage = colorBitmap.Clone();
            }

            capturedImages.Add(capturedImage);

            string fileName = $"photo_{DateTime.Now:yyyyMMddHHmmss}_{photoCount + 1}.png";
            string path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyPictures), "KinectPhotobooth", fileName);

            Directory.CreateDirectory(Path.GetDirectoryName(path));

            using (FileStream stream = new FileStream(path, FileMode.Create))
            {
                PngBitmapEncoder encoder = new PngBitmapEncoder();
                encoder.Frames.Add(BitmapFrame.Create(capturedImage));
                encoder.Save(stream);
            }

            photoCount++;

            if (photoCount < 3)
            {
                messageBlock.Text = $"Great pose! {3 - photoCount} more to go!";
                StartCountdown();
            }
            else
            {
                FinishPhotoSession();
            }
        }

        private BitmapSource ConvertBitmapToBitmapSource(System.Drawing.Bitmap bitmap)
        {
            var bitmapData = bitmap.LockBits(
                new System.Drawing.Rectangle(0, 0, bitmap.Width, bitmap.Height),
                System.Drawing.Imaging.ImageLockMode.ReadOnly, bitmap.PixelFormat);

            var bitmapSource = BitmapSource.Create(
                bitmapData.Width, bitmapData.Height,
                bitmap.HorizontalResolution, bitmap.VerticalResolution,
                PixelFormats.Bgr24, null,
                bitmapData.Scan0, bitmapData.Stride * bitmapData.Height, bitmapData.Stride);

            bitmap.UnlockBits(bitmapData);
            return bitmapSource;
        }

        private void StartCountdown()
        {
            countdownBlock.Text = "3";
            countdownTimer.Start();
        }

        private void FinishPhotoSession()
        {
            messageBlock.Text = "Thanks for coming! We'll send you your photos after the wedding!";
            countdownBlock.Text = "";
            DispatcherTimer resetTimer = new DispatcherTimer();
            resetTimer.Interval = TimeSpan.FromSeconds(5);
            resetTimer.Tick += (s, e) =>
            {
                resetTimer.Stop();
                ResetApplication();
            };
            resetTimer.Start();
        }

        private void ResetApplication()
        {
            photoCount = 0;
            capturedImages.Clear();
            ShowWelcomeMessage();
            goButton.Visibility = Visibility.Visible;
        }

        private void KinectSensor_ColorFrameReady(object sender, ColorImageFrameReadyEventArgs e)
        {
            using (ColorImageFrame colorFrame = e.OpenColorImageFrame())
            {
                if (colorFrame != null)
                {
                    colorFrame.CopyPixelDataTo(colorPixels);

                    colorBitmap.WritePixels(
                        new Int32Rect(0, 0, colorBitmap.PixelWidth, colorBitmap.PixelHeight),
                        colorPixels,
                        colorBitmap.PixelWidth * sizeof(int),
                        0);
                }
            }
        }

        private void KinectSensor_SkeletonFrameReady(object sender, SkeletonFrameReadyEventArgs e)
        {
            using (SkeletonFrame skeletonFrame = e.OpenSkeletonFrame())
            {
                if (skeletonFrame != null)
                {
                    if (skeletonData == null)
                    {
                        skeletonData = new Skeleton[skeletonFrame.SkeletonArrayLength];
                    }
                    skeletonFrame.CopySkeletonDataTo(skeletonData);
                    Skeleton skeleton = GetPrimarySkeleton(skeletonData);

                    if (skeleton != null)
                    {
                        TrackHand(skeleton);
                    }
                }
            }
        }

        private static Skeleton GetPrimarySkeleton(Skeleton[] skeletons)
        {
            Skeleton skeleton = null;
            if (skeletons != null)
            {
                for (int i = 0; i < skeletons.Length; i++)
                {
                    if (skeletons[i].TrackingState == SkeletonTrackingState.Tracked)
                    {
                        if (skeleton == null)
                        {
                            skeleton = skeletons[i];
                        }
                        else
                        {
                            if (skeletons[i].Position.Z < skeleton.Position.Z)
                            {
                                skeleton = skeletons[i];
                            }
                        }
                    }
                }
            }
            return skeleton;
        }

        private void TrackHand(Skeleton skeleton)
        {
            Joint rightHand = skeleton.Joints[JointType.HandRight];
            Point mappedPoint = SkeletonPointToScreen(rightHand.Position);

            // Debug information
            Console.WriteLine($"Hand position: X={mappedPoint.X}, Y={mappedPoint.Y}");
            Console.WriteLine($"Button position: Left={Canvas.GetLeft(goButton)}, Top={Canvas.GetTop(goButton)}");

            bool newIsHandOverButton = IsHandOverButton(mappedPoint);
            if (newIsHandOverButton != isHandOverButton)
            {
                isHandOverButton = newIsHandOverButton;
                UpdateButtonGlow();

                // Debug information
                Console.WriteLine($"Hand over button: {isHandOverButton}");
            }

            if (isHandOverButton)
            {
                if (handOverButtonStartTime == DateTime.MinValue)
                {
                    handOverButtonStartTime = DateTime.Now;
                }
                else if ((DateTime.Now - handOverButtonStartTime).TotalSeconds >= HandOverButtonThreshold)
                {
                    TakePhoto();
                    handOverButtonStartTime = DateTime.MinValue;
                }
            }
            else
            {
                handOverButtonStartTime = DateTime.MinValue;
            }
        }

        private Point SkeletonPointToScreen(SkeletonPoint skeletonPoint)
        {
            ColorImagePoint colorPoint = kinectSensor.CoordinateMapper.MapSkeletonPointToColorPoint(
                skeletonPoint, ColorImageFormat.RgbResolution640x480Fps30);

            // Scale the point to match the window size
            double scaleX = ActualWidth / kinectSensor.ColorStream.FrameWidth;
            double scaleY = ActualHeight / kinectSensor.ColorStream.FrameHeight;

            return new Point(colorPoint.X * scaleX, colorPoint.Y * scaleY);
        }

        private bool IsHandOverButton(Point handPosition)
        {
            double buttonLeft = Canvas.GetLeft(goButton);
            double buttonTop = Canvas.GetTop(goButton);
            double buttonRight = buttonLeft + goButton.ActualWidth;
            double buttonBottom = buttonTop + goButton.ActualHeight;

            // Increase the detection area by 50 pixels in each direction
            return handPosition.X >= buttonLeft - 50 && handPosition.X <= buttonRight + 50 &&
                   handPosition.Y >= buttonTop - 50 && handPosition.Y <= buttonBottom + 50;
        }

        private void UpdateButtonGlow()
        {
            if (isHandOverButton)
            {
                DoubleAnimation glowAnimation = new DoubleAnimation
                {
                    From = 0,
                    To = 1,
                    Duration = TimeSpan.FromSeconds(0.5),
                    AutoReverse = true,
                    RepeatBehavior = RepeatBehavior.Forever
                };
                goButton.BeginAnimation(OpacityProperty, glowAnimation);
            }
            else
            {
                goButton.BeginAnimation(OpacityProperty, null);
                goButton.Opacity = 1;
            }
        }

        private void TakePhoto()
        {
            goButton.Visibility = Visibility.Collapsed;
            messageBlock.Text = "Get ready for your photos!";
            StartCountdown();
        }

        protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
        {
            if (kinectSensor != null)
            {
                kinectSensor.Stop();
            }
            if (webcam != null && webcam.IsRunning)
            {
                webcam.SignalToStop();
                webcam.WaitForStop();
            }
            base.OnClosing(e);
        }
    }
}
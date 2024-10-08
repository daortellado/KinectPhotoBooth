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
using System.Windows.Media.Animation;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.Processing;
using SixLabors.ImageSharp.Drawing.Processing;
using SixLabors.ImageSharp.Formats.Png;
using System.Threading.Tasks;
using SixLabors.ImageSharp.PixelFormats;
using SixLabors.Fonts;
using WpfTextAlignment = System.Windows.TextAlignment;
using WpfHorizontalAlignment = System.Windows.HorizontalAlignment;
using WpfVerticalAlignment = System.Windows.VerticalAlignment;
using SixLabors.ImageSharp.Drawing;
using Path = System.IO.Path;

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

        private void SetupCountdownTimer()
        {
            countdownTimer = new DispatcherTimer();
            countdownTimer.Interval = TimeSpan.FromSeconds(1);
            countdownTimer.Tick += CountdownTimer_Tick;
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
            // Update layout
            UpdateLayout();

            // Center GO! button
            Canvas.SetLeft(goButton, (ActualWidth - goButton.ActualWidth) / 2);
            Canvas.SetTop(goButton, (ActualHeight - goButton.ActualHeight) / 2);

            // Position message block at the bottom
            Canvas.SetLeft(messageBlock, 50);
            Canvas.SetBottom(messageBlock, 50);
            messageBlock.Width = ActualWidth - 100;

            // Center countdown block
            Canvas.SetLeft(countdownBlock, (ActualWidth - countdownBlock.ActualWidth) / 2);
            Canvas.SetTop(countdownBlock, (ActualHeight - countdownBlock.ActualHeight) / 2);
        }

        private void TrackHand(Skeleton skeleton)
        {
            Joint rightHand = skeleton.Joints[JointType.HandRight];
            System.Windows.Point mappedPoint = SkeletonPointToScreen(rightHand.Position);

            bool newIsHandOverButton = IsHandOverButton(mappedPoint);
            if (newIsHandOverButton != isHandOverButton)
            {
                isHandOverButton = newIsHandOverButton;
                UpdateButtonGlow();
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
                Width = 300,
                Height = 300,
                FontSize = 72,
                FontWeight = FontWeights.Bold,
                Style = (Style)FindResource("WeddingButton")
            };
            mainCanvas.Children.Add(goButton);

            // Create message block
            messageBlock = new TextBlock
            {
                FontSize = 48,
                Foreground = System.Windows.Media.Brushes.White,
                TextAlignment = WpfTextAlignment.Center,
                HorizontalAlignment = WpfHorizontalAlignment.Stretch,
                TextWrapping = TextWrapping.Wrap,
                FontFamily = new System.Windows.Media.FontFamily("Segoe Script"),
                Effect = new DropShadowEffect
                {
                    Color = Colors.Black,
                    Direction = 320,
                    ShadowDepth = 5,
                    Opacity = 0.5
                }
            };
            mainCanvas.Children.Add(messageBlock);

            // Create countdown block
            countdownBlock = new TextBlock
            {
                FontSize = 300,
                Foreground = System.Windows.Media.Brushes.White,
                TextAlignment = WpfTextAlignment.Center,
                HorizontalAlignment = WpfHorizontalAlignment.Center,
                VerticalAlignment = WpfVerticalAlignment.Center,
                FontWeight = FontWeights.Bold,
                Effect = new DropShadowEffect
                {
                    Color = Colors.Black,
                    Direction = 320,
                    ShadowDepth = 5,
                    Opacity = 0.5
                },
                Visibility = Visibility.Collapsed
            };
            mainCanvas.Children.Add(countdownBlock);

            // Handle window resize
            SizeChanged += MainWindow_SizeChanged;
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

            System.Diagnostics.Debug.WriteLine($"Captured photo {capturedImages.Count}: {capturedImage.PixelWidth}x{capturedImage.PixelHeight}");

            photoCount++;

            if (photoCount < 3)
            {
                messageBlock.Text = $"Great pose! {3 - photoCount} more to go!";
                StartCountdown();
            }
            else
            {
                System.Diagnostics.Debug.WriteLine($"Total captured images: {capturedImages.Count}");
                FinishPhotoSession();
            }

            // Debug: Save each captured image separately
            SaveJpeg(capturedImage, Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyPictures), $"KinectPhotobooth\\debug_image_{photoCount}.jpg"), 100);
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

        private void DisplayCollage(BitmapSource collage)
        {
            System.Diagnostics.Debug.WriteLine($"Displaying collage: {collage.PixelWidth}x{collage.PixelHeight}");

            // Create a new Image control
            System.Windows.Controls.Image collageImage = new System.Windows.Controls.Image
            {
                Source = collage,
                Stretch = Stretch.Uniform
            };

            // Clear existing children and add the new Image
            mainCanvas.Children.Clear();
            mainCanvas.Children.Add(collageImage);

            // Ensure the Image fills the canvas
            collageImage.Width = mainCanvas.ActualWidth;
            collageImage.Height = mainCanvas.ActualHeight;

            System.Diagnostics.Debug.WriteLine($"Collage display size: {collageImage.Width}x{collageImage.Height}");

            // Add message and countdown blocks back
            mainCanvas.Children.Add(messageBlock);
            mainCanvas.Children.Add(countdownBlock);

            // Force layout update
            mainCanvas.UpdateLayout();
        }

        private void ShowPleaseWaitMessage()
        {
            messageBlock.Text = "Please wait, your photos are coming!";
            messageBlock.Visibility = Visibility.Visible;
            messageBlock.FontSize = 72; // Make the text larger
            messageBlock.Foreground = new SolidColorBrush(Colors.Red); // Make the text red for visibility

            // Force UI update
            Dispatcher.Invoke(() => { }, DispatcherPriority.Render);
        }

        private async void FinishPhotoSession()
        {
            System.Diagnostics.Debug.WriteLine($"FinishPhotoSession: Total captured images: {capturedImages.Count}");

            // Ensure we have exactly 3 photos
            if (capturedImages.Count != 3)
            {
                MessageBox.Show($"Error: Expected 3 photos, but got {capturedImages.Count}. Please try again.");
                ResetApplication();
                return;
            }

            // Show please wait message
            ShowPleaseWaitMessage();

            // Force UI update
            await Task.Delay(100);

            // Create 4x6 inch collage
            byte[] collageData = await CreateCollageWithImageSharp(capturedImages);
            BitmapSource collage = LoadCollage(collageData);

            if (collage == null)
            {
                MessageBox.Show("Failed to create collage. Please try again.");
                ResetApplication();
                return;
            }

            System.Diagnostics.Debug.WriteLine($"Collage created: {collage.PixelWidth}x{collage.PixelHeight}");

            // Save collage to network folder
            string networkFolderPath = @"\\WINDOWSXP\photoshare\photos";
            string fileName = $"collage_{DateTime.Now:yyyyMMddHHmmss}.png";
            string filePath = Path.Combine(networkFolderPath, fileName);

            try
            {
                SaveCollageAsPng(collage, filePath);
                System.Diagnostics.Debug.WriteLine($"Collage saved to: {filePath}");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to save collage to network folder: {ex.Message}\nSaving to local Pictures folder instead.");

                // Fallback to local Pictures folder if network save fails
                string localFolderPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyPictures), "KinectPhotobooth");
                Directory.CreateDirectory(localFolderPath);
                filePath = Path.Combine(localFolderPath, fileName);
                SaveCollageAsPng(collage, filePath);
                System.Diagnostics.Debug.WriteLine($"Collage saved to local folder: {filePath}");
            }

            messageBlock.Text = "Thanks for using the photobooth!";
            countdownBlock.Text = "";

            // Display the collage
            DisplayCollage(collage);

            DispatcherTimer resetTimer = new DispatcherTimer();
            resetTimer.Interval = TimeSpan.FromSeconds(5);
            resetTimer.Tick += (s, e) =>
            {
                resetTimer.Stop();
                ResetApplication();
            };
            resetTimer.Start();
        }
        private async Task<byte[]> CreateCollageWithImageSharp(List<BitmapSource> capturedImages)
        {
            // Define the size of the collage (4x6 inches at 300 DPI)
            int collageWidth = 1200; // 4 inches * 300 DPI
            int collageHeight = 1800; // 6 inches * 300 DPI

            // Calculate the size of each cell in the 2x3 grid
            int cellWidth = (collageWidth - 10) / 2; // Subtract 10 pixels for the center line
            int cellHeight = collageHeight / 3;

            using (var collage = new SixLabors.ImageSharp.Image<Rgba32>(collageWidth, collageHeight))
            {
                collage.Mutate(ctx =>
                {
                    ctx.BackgroundColor(SixLabors.ImageSharp.Color.White);

                    // Draw the center line
                    var centerLine = new SixLabors.ImageSharp.Drawing.PathBuilder()
                        .AddLine(new SixLabors.ImageSharp.PointF(collageWidth / 2, 0), new SixLabors.ImageSharp.PointF(collageWidth / 2, collageHeight))
                        .Build();
                    ctx.Draw(SixLabors.ImageSharp.Color.Black, 2f, centerLine);

                    for (int i = 0; i < 3; i++)
                    {
                        var sourceImage = capturedImages[i];

                        // Convert BitmapSource to byte array
                        byte[] imageData;
                        using (var stream = new MemoryStream())
                        {
                            var encoder = new PngBitmapEncoder();
                            encoder.Frames.Add(BitmapFrame.Create(sourceImage));
                            encoder.Save(stream);
                            imageData = stream.ToArray();
                        }

                        using (var image = SixLabors.ImageSharp.Image.Load(imageData))
                        {
                            // Calculate scaling factors
                            float scaleX = (float)cellWidth / image.Width;
                            float scaleY = (float)cellHeight / image.Height;
                            float scale = Math.Min(scaleX, scaleY);

                            // Resize the image
                            image.Mutate(x => x.Resize((int)(image.Width * scale), (int)(image.Height * scale)));

                            // Calculate the position to center the image within the cell
                            int offsetX = (cellWidth - image.Width) / 2;
                            int offsetY = (cellHeight - image.Height) / 2;

                            // Add the image to both columns of the current row
                            for (int col = 0; col < 2; col++)
                            {
                                int x = col * (cellWidth + 10) + offsetX; // Add 10 pixels for the center line
                                int y = i * cellHeight + offsetY;

                                ctx.DrawImage(image, new SixLabors.ImageSharp.Point(x, y), 1f);
                            }
                        }
                    }

                    // Add text between rows
                    var font = new SixLabors.Fonts.Font(SixLabors.Fonts.SystemFonts.Get("Segoe Script"), 48);

                    // Function to draw centered text in a column
                    Action<string, float, float> drawCenteredTextInColumn = (text, x, y) =>
                    {
                        // Estimate text width (this is a rough estimate and may need adjustment)
                        float estimatedWidth = text.Length * font.Size * 0.6f;
                        float textX = x + (cellWidth - estimatedWidth) / 2;
                        ctx.DrawText(text, font, SixLabors.ImageSharp.Color.Black, new SixLabors.ImageSharp.PointF(textX, y - font.Size / 2));
                    };

                    // Draw "Katelyn + Damian" in both columns
                    drawCenteredTextInColumn("Katelyn + Damian", 0, cellHeight);
                    drawCenteredTextInColumn("Katelyn + Damian", cellWidth + 10, cellHeight);

                    // Draw "Oct. 12, 2024" in both columns
                    drawCenteredTextInColumn("Oct. 12, 2024", 0, 2 * cellHeight);
                    drawCenteredTextInColumn("Oct. 12, 2024", cellWidth + 10, 2 * cellHeight);

                    // Add a decorative border
                    var border = new SixLabors.ImageSharp.Drawing.PathBuilder()
                        .AddLine(new SixLabors.ImageSharp.PointF(0, 0), new SixLabors.ImageSharp.PointF(collageWidth - 1, 0))
                        .AddLine(new SixLabors.ImageSharp.PointF(collageWidth - 1, 0), new SixLabors.ImageSharp.PointF(collageWidth - 1, collageHeight - 1))
                        .AddLine(new SixLabors.ImageSharp.PointF(collageWidth - 1, collageHeight - 1), new SixLabors.ImageSharp.PointF(0, collageHeight - 1))
                        .AddLine(new SixLabors.ImageSharp.PointF(0, collageHeight - 1), new SixLabors.ImageSharp.PointF(0, 0))
                        .Build();
                    ctx.Draw(SixLabors.ImageSharp.Color.Gold, 5f, border);
                });

                // Save the collage to a byte array
                using (var ms = new MemoryStream())
                {
                    await collage.SaveAsPngAsync(ms);
                    return ms.ToArray();
                }
            }
        }

        private BitmapSource LoadCollage(byte[] collageData)
        {
            using (var ms = new MemoryStream(collageData))
            {
                var decoder = new PngBitmapDecoder(ms, BitmapCreateOptions.PreservePixelFormat, BitmapCacheOption.OnLoad);
                return decoder.Frames[0];
            }
        }

        private void SaveCollageAsPng(BitmapSource collage, string filePath)
        {
            System.Diagnostics.Debug.WriteLine($"Saving collage: {collage.PixelWidth}x{collage.PixelHeight}");

            try
            {
                using (FileStream stream = new FileStream(filePath, FileMode.Create))
                {
                    PngBitmapEncoder encoder = new PngBitmapEncoder();
                    encoder.Frames.Add(BitmapFrame.Create(collage));
                    encoder.Save(stream);
                }

                // Verify the saved file
                using (FileStream verifyStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    BitmapDecoder decoder = BitmapDecoder.Create(verifyStream, BitmapCreateOptions.PreservePixelFormat, BitmapCacheOption.Default);
                    BitmapSource savedImage = decoder.Frames[0];
                    System.Diagnostics.Debug.WriteLine($"Verified saved collage: {savedImage.PixelWidth}x{savedImage.PixelHeight}");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error saving collage: {ex.Message}");
                MessageBox.Show($"Error saving collage: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void SaveJpeg(BitmapSource bitmap, string filePath, int quality)
        {
            using (FileStream stream = new FileStream(filePath, FileMode.Create))
            {
                JpegBitmapEncoder encoder = new JpegBitmapEncoder();
                encoder.QualityLevel = quality;
                encoder.Frames.Add(BitmapFrame.Create(bitmap));
                encoder.Save(stream);
            }
        }

        private void ResetApplication()
        {
            photoCount = 0;
            capturedImages.Clear();

            // Clear the canvas and re-add necessary elements
            mainCanvas.Children.Clear();
            mainCanvas.Children.Add(goButton);
            mainCanvas.Children.Add(messageBlock);
            mainCanvas.Children.Add(countdownBlock);

            goButton.Visibility = Visibility.Visible;
            messageBlock.Text = "Hold your hand over the Start button to begin!";
            messageBlock.Visibility = Visibility.Visible;
            countdownBlock.Visibility = Visibility.Collapsed;

            // Reset the background
            mainCanvas.Background = null;
        }

        private void TakePhoto()
        {
            goButton.Visibility = Visibility.Collapsed;
            messageBlock.Text = "Get ready for your photos!";
            countdownBlock.Visibility = Visibility.Visible;
            StartCountdown();
        }

        private void StartCountdown()
        {
            countdownBlock.Text = "3";
            countdownTimer.Start();
        }

        private System.Windows.Point SkeletonPointToScreen(SkeletonPoint skeletonPoint)
        {
            ColorImagePoint colorPoint = kinectSensor.CoordinateMapper.MapSkeletonPointToColorPoint(
                skeletonPoint, ColorImageFormat.RgbResolution640x480Fps30);

            // Scale the point to match the window size
            double scaleX = ActualWidth / kinectSensor.ColorStream.FrameWidth;
            double scaleY = ActualHeight / kinectSensor.ColorStream.FrameHeight;

            return new System.Windows.Point(colorPoint.X * scaleX, colorPoint.Y * scaleY);
        }

        private bool IsHandOverButton(System.Windows.Point handPosition)
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
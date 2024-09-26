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
using System.Linq;

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

        // New variables for party selection
        private List<string> partyNames;
        private int currentPage = 0;
        private const int partiesPerPage = 9; // 3x3 grid
        private string selectedParty;

        // UI elements for party selection
        private Grid partySelectionGrid;
        private Button nextPageButton;
        private Button prevPageButton;
        private TextBlock titleBlock;

        public MainWindow()
        {
            InitializeComponent();
            InitializeKinect();
            InitializeWebcam();
            LoadPartyNames();
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

        private void LoadPartyNames()
        {
            string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "party_names.csv");
            partyNames = File.ReadAllLines(path).Skip(1).ToList(); // Skip header
        }

        private const int ButtonsPerPage = 16; // 4x4 grid
        private Button hoveredButton;

        private void CreateUI()
        {
            // Create title block
            titleBlock = new TextBlock
            {
                Text = "Select Your Party",
                FontSize = 72,
                Foreground = Brushes.White,
                TextAlignment = TextAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Stretch,
                VerticalAlignment = VerticalAlignment.Top,
                Margin = new Thickness(0, 20, 0, 0)
            };
            mainCanvas.Children.Add(titleBlock);

            // Create party selection grid
            partySelectionGrid = new Grid
            {
                HorizontalAlignment = HorizontalAlignment.Stretch,
                VerticalAlignment = VerticalAlignment.Stretch,
                Margin = new Thickness(0, 100, 0, 100) // Add top and bottom margin
            };
            mainCanvas.Children.Add(partySelectionGrid);

            // Create navigation buttons
            nextPageButton = CreateNavigationButton("Next", 20);
            prevPageButton = CreateNavigationButton("Prev", 20);
            mainCanvas.Children.Add(nextPageButton);
            mainCanvas.Children.Add(prevPageButton);

            // Create GO! button (initially hidden)
            goButton = new Button
            {
                Content = "Start",
                Width = 300,
                Height = 300,
                FontSize = 72,
                FontWeight = FontWeights.Bold,
                Style = (Style)FindResource("WeddingButton"),
                Visibility = Visibility.Collapsed
            };
            mainCanvas.Children.Add(goButton);

            // Create message block
            messageBlock = new TextBlock
            {
                FontSize = 48,
                Foreground = Brushes.White,
                TextAlignment = TextAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Stretch,
                TextWrapping = TextWrapping.Wrap,
                FontFamily = new FontFamily("Segoe Script"),
                Effect = new DropShadowEffect
                {
                    Color = Colors.Black,
                    Direction = 320,
                    ShadowDepth = 5,
                    Opacity = 0.5
                },
                Visibility = Visibility.Collapsed
            };
            mainCanvas.Children.Add(messageBlock);

            // Create countdown block
            countdownBlock = new TextBlock
            {
                FontSize = 300,
                Foreground = Brushes.White,
                TextAlignment = TextAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center,
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

            UpdatePartySelectionUI();

            // Handle window resize
            SizeChanged += MainWindow_SizeChanged;
        }

        private void UpdatePartySelectionUI()
        {
            partySelectionGrid.Children.Clear();
            partySelectionGrid.RowDefinitions.Clear();
            partySelectionGrid.ColumnDefinitions.Clear();

            // Create grid with 4 columns
            for (int i = 0; i < 4; i++)
            {
                partySelectionGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            }

            // Add buttons for each party
            int startIndex = currentPage * ButtonsPerPage;
            for (int i = 0; i < ButtonsPerPage; i++)
            {
                if (startIndex + i >= partyNames.Count) break;

                // Create a button and set its properties
                var partyButton = new Button
                {
                    Content = partyNames[startIndex + i],
                    FontSize = 48,
                    Margin = new Thickness(10),
                    Style = (Style)FindResource("WeddingButton"),
                    HorizontalAlignment = HorizontalAlignment.Stretch,
                    VerticalAlignment = VerticalAlignment.Stretch,
                    Background = Brushes.Transparent // Set initial background to transparent
                };

                // Add the button to the grid
                partySelectionGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
                Grid.SetRow(partyButton, i / 4);
                Grid.SetColumn(partyButton, i % 4);
                partySelectionGrid.Children.Add(partyButton);

                // Add hover event handler
                partyButton.MouseEnter += PartyButton_MouseEnter;
                partyButton.MouseLeave += PartyButton_MouseLeave;
            }

            // Update visibility of navigation buttons
            prevPageButton.Visibility = currentPage > 0 ? Visibility.Visible : Visibility.Collapsed;
            nextPageButton.Visibility = (currentPage + 1) * ButtonsPerPage < partyNames.Count ? Visibility.Visible : Visibility.Collapsed;
        }

        private void PartyButton_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            Button button = (Button)sender;
            button.Background = Brushes.LightBlue;
        }

        private void PartyButton_MouseLeave(object sender, System.Windows.Input.MouseEventArgs e)
        {
            Button button = (Button)sender;
            button.Background = Brushes.Transparent;
        }


        private Button CreateNavigationButton(string content, double margin)
        {
            var button = new Button
            {
                Content = content,
                Width = 100,
                Height = 100,
                FontSize = 24,
                Background = Brushes.Blue,
                Foreground = Brushes.White,
                VerticalAlignment = VerticalAlignment.Center
            };
            return button;
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

            // Position navigation buttons
            Canvas.SetLeft(prevPageButton, 20);
            Canvas.SetRight(nextPageButton, 20);
            Canvas.SetTop(prevPageButton, (ActualHeight - prevPageButton.ActualHeight) / 2);
            Canvas.SetTop(nextPageButton, (ActualHeight - nextPageButton.ActualHeight) / 2);

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

        private void FinishPhotoSession()
        {
            // Create collage
            var collage = CreateCollage();

            // Save collage
            string folderPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyPictures), "KinectPhotobooth", selectedParty);
            Directory.CreateDirectory(folderPath);
            string fileName = $"collage_{DateTime.Now:yyyyMMddHHmmss}.png";
            string filePath = Path.Combine(folderPath, fileName);

            using (FileStream stream = new FileStream(filePath, FileMode.Create))
            {
                PngBitmapEncoder encoder = new PngBitmapEncoder();
                encoder.Frames.Add(BitmapFrame.Create(collage));
                encoder.Save(stream);
            }

            messageBlock.Text = "Thanks for coming!";
            countdownBlock.Text = "";

            // Show logo
            System.Windows.Controls.Image logoImage = new System.Windows.Controls.Image
            {
                Source = new BitmapImage(new Uri("path/to/your/logo.png", UriKind.Relative)),
                Width = 200,
                Height = 200
            };
            Canvas.SetLeft(logoImage, (ActualWidth - logoImage.Width) / 2);
            Canvas.SetTop(logoImage, (ActualHeight - logoImage.Height) / 2);
            mainCanvas.Children.Add(logoImage);

            DispatcherTimer resetTimer = new DispatcherTimer();
            resetTimer.Interval = TimeSpan.FromSeconds(5);
            resetTimer.Tick += (s, e) =>
            {
                resetTimer.Stop();
                mainCanvas.Children.Remove(logoImage);
                ResetApplication();
            };
            resetTimer.Start();
        }
        private WriteableBitmap CreateCollage()
        {
            int width = capturedImages[0].PixelWidth;
            int height = capturedImages[0].PixelHeight * 3 + 100; // Extra space for text
            var collage = new WriteableBitmap(width, height, 96, 96, PixelFormats.Bgr32, null);

            // Create a DrawingVisual
            DrawingVisual drawingVisual = new DrawingVisual();
            using (DrawingContext drawingContext = drawingVisual.RenderOpen())
            {
                for (int i = 0; i < 3; i++)
                {
                    drawingContext.DrawImage(capturedImages[i], new Rect(0, i * capturedImages[i].PixelHeight, width, capturedImages[i].PixelHeight));
                }

                // Add text
                FormattedText formattedText = new FormattedText(
                    $"Katelyn and Damian - Oct. 12, '24",
                    System.Globalization.CultureInfo.CurrentCulture,
                    FlowDirection.LeftToRight,
                    new Typeface("Arial"),
                    24,
                    Brushes.White,
                    VisualTreeHelper.GetDpi(this).PixelsPerDip);

                drawingContext.DrawText(formattedText, new Point((width - formattedText.Width) / 2, height - 50));
            }

            // Render the DrawingVisual to the WriteableBitmap
            RenderTargetBitmap renderBitmap = new RenderTargetBitmap(width, height, 96, 96, PixelFormats.Pbgra32);
            renderBitmap.Render(drawingVisual);

            // Copy the rendered bitmap to the collage WriteableBitmap
            collage.Lock();
            renderBitmap.CopyPixels(new Int32Rect(0, 0, width, height), collage.BackBuffer, collage.BackBufferStride * height, collage.BackBufferStride);
            collage.AddDirtyRect(new Int32Rect(0, 0, width, height));
            collage.Unlock();

            return collage;
        }

        private void ResetApplication()
        {
            photoCount = 0;
            capturedImages.Clear();
            selectedParty = null;
            currentPage = 0;
            UpdatePartySelectionUI();
            titleBlock.Visibility = Visibility.Visible;
            partySelectionGrid.Visibility = Visibility.Visible;
            goButton.Visibility = Visibility.Collapsed;
            messageBlock.Visibility = Visibility.Collapsed;
            countdownBlock.Visibility = Visibility.Collapsed;
        }

        private void TrackHand(Skeleton skeleton)
        {
            Joint rightHand = skeleton.Joints[JointType.HandRight];
            Point mappedPoint = SkeletonPointToScreen(rightHand.Position);

            if (selectedParty == null)
            {
                // Check for hovering over party selection buttons
                Button newHoveredButton = null;
                foreach (Button button in partySelectionGrid.Children.OfType<Button>())
                {
                    if (IsPointOverElement(mappedPoint, button))
                    {
                        newHoveredButton = button;
                        break;
                    }
                }

                // Check for hovering over navigation buttons
                if (newHoveredButton == null)
                {
                    if (IsPointOverElement(mappedPoint, nextPageButton) && nextPageButton.Visibility == Visibility.Visible)
                    {
                        newHoveredButton = nextPageButton;
                    }
                    else if (IsPointOverElement(mappedPoint, prevPageButton) && prevPageButton.Visibility == Visibility.Visible)
                    {
                        newHoveredButton = prevPageButton;
                    }
                }

                // Apply visual feedback for hovered button
                if (newHoveredButton != hoveredButton)
                {
                    if (hoveredButton != null)
                    {
                        hoveredButton.Background = Brushes.Transparent; // Reset background on previous button
                    }
                    hoveredButton = newHoveredButton;
                    if (hoveredButton != null)
                    {
                        hoveredButton.Background = Brushes.LightBlue; // Highlight the hovered button
                    }
                }

                // Handle button clicks
                if (hoveredButton != null)
                {
                    if (handOverButtonStartTime == DateTime.MinValue)
                    {
                        handOverButtonStartTime = DateTime.Now;
                    }
                    else if ((DateTime.Now - handOverButtonStartTime).TotalSeconds >= HandOverButtonThreshold)
                    {
                        if (hoveredButton == nextPageButton)
                        {
                            currentPage++;
                            UpdatePartySelectionUI();
                        }
                        else if (hoveredButton == prevPageButton)
                        {
                            currentPage--;
                            UpdatePartySelectionUI();
                        }
                        else
                        {
                            selectedParty = hoveredButton.Content.ToString();
                            StartPhotoSession();
                        }
                        handOverButtonStartTime = DateTime.MinValue;
                    }
                }
                else
                {
                    handOverButtonStartTime = DateTime.MinValue;
                }

                // Check for navigation buttons
                if (IsPointOverElement(mappedPoint, nextPageButton) && nextPageButton.Visibility == Visibility.Visible)
                {
                    if (handOverButtonStartTime == DateTime.MinValue)
                    {
                        handOverButtonStartTime = DateTime.Now;
                    }
                    else if ((DateTime.Now - handOverButtonStartTime).TotalSeconds >= HandOverButtonThreshold)
                    {
                        currentPage++;
                        UpdatePartySelectionUI();
                        handOverButtonStartTime = DateTime.MinValue;
                    }
                }
                else if (IsPointOverElement(mappedPoint, prevPageButton) && prevPageButton.Visibility == Visibility.Visible)
                {
                    if (handOverButtonStartTime == DateTime.MinValue)
                    {
                        handOverButtonStartTime = DateTime.Now;
                    }
                    else if ((DateTime.Now - handOverButtonStartTime).TotalSeconds >= HandOverButtonThreshold)
                    {
                        currentPage--;
                        UpdatePartySelectionUI();
                        handOverButtonStartTime = DateTime.MinValue;
                    }
                }
            }
            else
            {
                // Existing logic for "Go!" button
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
        }

        private void SelectParty(Point handPosition)
        {
            foreach (Button button in partySelectionGrid.Children.OfType<Button>())
            {
                if (IsPointOverElement(handPosition, button))
                {
                    selectedParty = button.Content.ToString();
                    StartPhotoSession();
                    return;
                }
            }

            if (IsPointOverElement(handPosition, nextPageButton) && nextPageButton.Visibility == Visibility.Visible)
            {
                currentPage++;
                UpdatePartySelectionUI();
            }
            else if (IsPointOverElement(handPosition, prevPageButton) && prevPageButton.Visibility == Visibility.Visible)
            {
                currentPage--;
                UpdatePartySelectionUI();
            }
        }

        private bool IsPointOverElement(Point point, FrameworkElement element)
        {
            Point relativePoint = element.TransformToAncestor(mainCanvas).Transform(new Point(0, 0));
            return point.X >= relativePoint.X && point.X <= relativePoint.X + element.ActualWidth &&
                   point.Y >= relativePoint.Y && point.Y <= relativePoint.Y + element.ActualHeight;
        }

        private void StartPhotoSession()
        {
            titleBlock.Visibility = Visibility.Collapsed;
            partySelectionGrid.Visibility = Visibility.Collapsed;
            nextPageButton.Visibility = Visibility.Collapsed;
            prevPageButton.Visibility = Visibility.Collapsed;

            goButton.Visibility = Visibility.Visible;
            messageBlock.Visibility = Visibility.Visible;
            messageBlock.Text = $"Selected Party: {selectedParty}\nHold your hand over the Start button to begin!";
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
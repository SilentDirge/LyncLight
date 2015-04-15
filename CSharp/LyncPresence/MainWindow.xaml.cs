/*=====================================================================
  This file is part of the Microsoft Unified Communications Code Samples.

  Copyright (C) 2012 Microsoft Corporation.  All rights reserved.

This source code is intended only as a supplement to Microsoft
Development Tools and/or on-line documentation.  See these other
materials for detailed information regarding Microsoft code samples.

THIS CODE AND INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY
KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE
IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
PARTICULAR PURPOSE.
=====================================================================*/
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using Microsoft.Lync.Model;
using System.IO.Ports;
using System.Windows.Forms;

namespace LyncPresence
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        #region Fields
        // Current dispatcher reference for changes in the user interface.
        private Dispatcher dispatcher;
        private LyncClient lyncClient;

        private SerialPort serialPort;

        static Timer sleepTimer;

        bool isLightshowModeActive = false;
        #endregion

        public MainWindow()
        {
            InitializeComponent();

            Title = "Initializing...";

            //Save the current dispatcher to use it for changes in the user interface.
            dispatcher = Dispatcher.CurrentDispatcher;

            // Get a list of serial port names. 
            string[] ports = SerialPort.GetPortNames();

            Console.WriteLine("The following serial ports were found:");

            // Display each port name to the console. 
            foreach (string port in ports)
            {
                Console.WriteLine(port);
            }

            try
            {
                serialPort = new SerialPort("COM8", 9600, Parity.None, 8, StopBits.One);
                serialPort.Open();
                Console.WriteLine("SerialPort Opened Successfully");

                ApplyColor(MyColor.Blue);

                DateTime timeNow = DateTime.Now;
                // started within light show period?
                if (timeNow.Hour >= 19 || timeNow.Hour < 8 || timeNow.DayOfWeek == DayOfWeek.Saturday || timeNow.DayOfWeek == DayOfWeek.Sunday)
                {
                    // go straight into the light show
                    isLightshowModeActive = true;
                    ApplyColor(MyColor.LightshowMode);

                    DateTime nextTime;

                    if (timeNow.DayOfWeek == DayOfWeek.Saturday)
                    {
                        DateTime monday = DateTime.Now + new TimeSpan(2, 0, 0, 0);
                        nextTime = new DateTime(monday.Year, monday.Month, monday.Day, 8, 0, 0);
                    }
                    else if (timeNow.DayOfWeek == DayOfWeek.Sunday)
                    {
                        DateTime monday = DateTime.Now + new TimeSpan(1, 0, 0, 0);
                        nextTime = new DateTime(monday.Year, monday.Month, monday.Day, 8, 0, 0);
                    }
                    else if (timeNow.Hour >= 19)
                    {
                        DateTime tomorrow = DateTime.Now + new TimeSpan(1, 0, 0, 0);
                        nextTime = new DateTime(tomorrow.Year, tomorrow.Month, tomorrow.Day, 8, 0, 0);
                    }
                    else
                    {
                        DateTime now = DateTime.Now;
                        nextTime = new DateTime(now.Year, now.Month, now.Day, 8, 0, 0);
                    }

                    SetupSleepTimer(nextTime);
                }
                else
                {
                    isLightshowModeActive = false;
                    SetupSleepTimer(GetNextTime(isLightshowModeActive));
                }

                Title = "Lync Presence: Connected";
            }
            catch (SystemException systemException)
            {
                serialPort = null;
                Console.WriteLine("SerialPort Error: " + systemException);
                Title = "Lync Presence: *Disconnected*";
            }
        }

        public void SetupSleepTimer(DateTime futureTime)
        {
            sleepTimer = new Timer();
            sleepTimer.Tick += (sender, eventArgs) =>
            {
                sleepTimer.Stop();

                // already in light show so disable it and go to the current status light
                if (isLightshowModeActive)
                {
                    isLightshowModeActive = false;
                    SetupSleepTimer(GetNextTime(isLightshowModeActive));
                    dispatcher.BeginInvoke(new Action(SetAvailability));
                }
                else
                {
                    isLightshowModeActive = true;
                    SetupSleepTimer(GetNextTime(isLightshowModeActive));
                    ApplyColor(MyColor.LightshowMode);
                }
            };

            DateTime timeNow = DateTime.Now;
            TimeSpan timeDif = futureTime - timeNow;

            sleepTimer.Interval = (int)timeDif.TotalMilliseconds;
            sleepTimer.Start();

            if (isLightshowModeActive)
            {
                Console.WriteLine("Time until light show ends: " + timeDif.ToString());
            }
            else
            {
                Console.WriteLine("Time until light show begins: " + timeDif.ToString());
            }
        }

        private DateTime GetNextTime(bool lightshowActive)
        {
            DateTime timeNow = DateTime.Now;

            if (timeNow.DayOfWeek == DayOfWeek.Saturday)
            {
                DateTime monday = DateTime.Now + new TimeSpan(2, 0, 0, 0);
                return new DateTime(monday.Year, monday.Month, monday.Day, 8, 0, 0);
            }
            else if (timeNow.DayOfWeek == DayOfWeek.Sunday)
            {
                DateTime monday = DateTime.Now + new TimeSpan(1, 0, 0, 0);
                return new DateTime(monday.Year, monday.Month, monday.Day, 8, 0, 0);
            }

            // if in lightshow, check again (and stop) after x hours from now
            if (lightshowActive)
            {
                // turn off at 8 am the following day
                DateTime tomorrow = DateTime.Now + new TimeSpan(1, 0, 0, 0);
                return new DateTime(tomorrow.Year, tomorrow.Month, tomorrow.Day, 8, 0, 0);
                // TESTING:
                //return DateTime.Now + new TimeSpan(0, 0, 1, 0, 0);
            }
            else
            {
                // start at 7pm
                return new DateTime(timeNow.Year, timeNow.Month, timeNow.Day, 19, 0, 0);
                // TESTING:
                //return DateTime.Now + new TimeSpan(0, 0, 1, 0, 0);
            }
        }

        #region Handlers for user interface controls events
        /// <summary>
        /// Handler for the Loaded event of the Window.
        /// Used to initialize the values shown in the user interface (e.g. availability values), get the Lync client instance
        /// and start listening for events of changes in the client state.
        /// </summary>
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            //Add the availability values to the ComboBox
            availabilityComboBox.Items.Add(ContactAvailability.Free);
            availabilityComboBox.Items.Add(ContactAvailability.Busy);
            availabilityComboBox.Items.Add(ContactAvailability.DoNotDisturb);
            availabilityComboBox.Items.Add(ContactAvailability.Away);

            //Listen for events of changes in the state of the client
            try
            {
                lyncClient = LyncClient.GetClient();
            }
            catch (ClientNotFoundException clientNotFoundException)
            {
                Console.WriteLine(clientNotFoundException);
                return;
            }
            catch (NotStartedByUserException notStartedByUserException)
            {
                Console.Out.WriteLine(notStartedByUserException);
                return;
            }
            catch (LyncClientException lyncClientException)
            {
                Console.Out.WriteLine(lyncClientException);
                return;
            }
            catch (SystemException systemException)
            {
                if (IsLyncException(systemException))
                {
                    // Log the exception thrown by the Lync Model API.
                    Console.WriteLine("Error: " + systemException);
                    return;
                }
                else
                {
                    // Rethrow the SystemException which did not come from the Lync Model API.
                    throw;
                }
            }

            lyncClient.StateChanged +=
                new EventHandler<ClientStateChangedEventArgs>(Client_StateChanged);

            //Update the user interface
            UpdateUserInterface(lyncClient.State);
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (serialPort != null)
            {
                isLightshowModeActive = false;
                ApplyColor(MyColor.Magenta);
                serialPort.Close();
            }
        }

        /// <summary>
        /// Handler for the SelectionChanged event of the Availability ComboBox. Used to publish the selected availability value in Lync
        /// </summary>
        private void AvailabilityComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            /*
            //Add the availability to the contact information items to be published
            Dictionary<PublishableContactInformationType, object> newInformation =
                new Dictionary<PublishableContactInformationType, object>();
            newInformation.Add(PublishableContactInformationType.Availability, availabilityComboBox.SelectedItem);

            //Publish the new availability value
            try
            {
                lyncClient.Self.BeginPublishContactInformation(newInformation, PublishContactInformationCallback, null);
            }
            catch (LyncClientException lyncClientException)
            {
                Console.WriteLine(lyncClientException);
            }
            catch (SystemException systemException)
            {
                if (IsLyncException(systemException))
                {
                    // Log the exception thrown by the Lync Model API.
                    Console.WriteLine("Error: " + systemException);
                }
                else
                {
                    // Rethrow the SystemException which did not come from the Lync Model API.
                    throw;
                }
            }
             */

        }

        /// <summary>
        /// Handler for the Click event of the Note Button. Used to publish a new personal note value in Lync
        /// </summary>
        private void SetNoteButton_Click(object sender, RoutedEventArgs e)
        {
            //Add the personal note to the contact information items to be published
            Dictionary<PublishableContactInformationType, object> newInformation =
                new Dictionary<PublishableContactInformationType, object>();
            newInformation.Add(PublishableContactInformationType.PersonalNote, personalNoteTextBox.Text);

            //Publish the new personal note value
            try
            {
                lyncClient.Self.BeginPublishContactInformation(newInformation, PublishContactInformationCallback, null);
            }
            catch (LyncClientException lyncClientException)
            {
                Console.WriteLine(lyncClientException);
            }
            catch (SystemException systemException)
            {
                if (IsLyncException(systemException))
                {
                    // Log the exception thrown by the Lync Model API.
                    Console.WriteLine("Error: " + systemException);
                }
                else
                {
                    // Rethrow the SystemException which did not come from the Lync Model API.
                    throw;
                }
            }
        }

        /// <summary>
        /// Handler for the Click event of the SignInOut Button. Used to sign in or out Lync depending on the current client state.
        /// </summary>
        private void SignInOutButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (lyncClient.State == ClientState.SignedIn)
                {
                    //Sign out If the current client state is Signed In
                    lyncClient.BeginSignOut(SignOutCallback, null);
                }
                else if (lyncClient.State == ClientState.SignedOut)
                {
                    //Sign in If the current client state is Signed Out
                    lyncClient.BeginSignIn(null, null, null, SignInCallback, null);
                }
            }
            catch (LyncClientException lyncClientException)
            {
                Console.WriteLine(lyncClientException);
            }
            catch (SystemException systemException)
            {
                if (IsLyncException(systemException))
                {
                    // Log the exception thrown by the Lync Model API.
                    Console.WriteLine("Error: " + systemException);
                }
                else
                {
                    // Rethrow the SystemException which did not come from the Lync Model API.
                    throw;
                }
            }

        }
        #endregion

        #region Handlers for Lync events
        /// <summary>
        /// Handler for the ContactInformationChanged event of the contact. Used to update the contact's information in the user interface.
        /// </summary>
        private void SelfContact_ContactInformationChanged(object sender, ContactInformationChangedEventArgs e)
        {
            //Only update the contact information in the user interface if the client is signed in.
            //Ignore other states including transitions (e.g. signing in or out).
            if (lyncClient.State == ClientState.SignedIn)
            {
                //Get from Lync only the contact information that changed.

                if (e.ChangedContactInformation.Contains(ContactInformationType.DisplayName))
                {
                    //Use the current dispatcher to update the contact's name in the user interface.
                    dispatcher.BeginInvoke(new Action(SetName));
                }
                if (e.ChangedContactInformation.Contains(ContactInformationType.Availability))
                {
                    //Use the current dispatcher to update the contact's availability in the user interface.
                    dispatcher.BeginInvoke(new Action(SetAvailability));
                }
                if (e.ChangedContactInformation.Contains(ContactInformationType.PersonalNote))
                {
                    //Use the current dispatcher to update the contact's personal note in the user interface.
                    dispatcher.BeginInvoke(new Action(SetPersonalNote));
                }
                if (e.ChangedContactInformation.Contains(ContactInformationType.Photo))
                {
                    //Use the current dispatcher to update the contact's photo in the user interface.
                    dispatcher.BeginInvoke(new Action(SetContactPhoto));
                }
            }
        }

        /// <summary>
        /// Handler for the StateChanged event of the contact. Used to update the user interface with the new client state.
        /// </summary>
        private void Client_StateChanged(object sender, ClientStateChangedEventArgs e)
        {
            //Use the current dispatcher to update the user interface with the new client state.
            dispatcher.BeginInvoke(new Action<ClientState>(UpdateUserInterface), e.NewState);
        }
        #endregion

        #region Callbacks
        /// <summary>
        /// Callback invoked when LyncClient.BeginSignIn is completed
        /// </summary>
        /// <param name="result">The status of the asynchronous operation</param>
        private void SignInCallback(IAsyncResult result)
        {
            try
            {
                lyncClient.EndSignIn(result);
            }
            catch (LyncClientException e)
            {
                Console.WriteLine(e);
            }
            catch (SystemException systemException)
            {
                if (IsLyncException(systemException))
                {
                    // Log the exception thrown by the Lync Model API.
                    Console.WriteLine("Error: " + systemException);
                }
                else
                {
                    // Rethrow the SystemException which did not come from the Lync Model API.
                    throw;
                }
            }

        }

        /// <summary>
        /// Callback invoked when LyncClient.BeginSignOut is completed
        /// </summary>
        /// <param name="result">The status of the asynchronous operation</param>
        private void SignOutCallback(IAsyncResult result)
        {
            try
            {
                lyncClient.EndSignOut(result);
            }
            catch (LyncClientException e)
            {
                Console.WriteLine(e);
            }
            catch (SystemException systemException)
            {
                if (IsLyncException(systemException))
                {
                    // Log the exception thrown by the Lync Model API.
                    Console.WriteLine("Error: " + systemException);
                }
                else
                {
                    // Rethrow the SystemException which did not come from the Lync Model API.
                    throw;
                }
            }

        }

        /// <summary>
        /// Callback invoked when Self.BeginPublishContactInformation is completed
        /// </summary>
        /// <param name="result">The status of the asynchronous operation</param>
        private void PublishContactInformationCallback(IAsyncResult result)
        {
            lyncClient.Self.EndPublishContactInformation(result);
        }
        #endregion

        /// <summary>
        /// Updates the user interface
        /// </summary>
        /// <param name="currentState"></param>
        private void UpdateUserInterface(ClientState currentState)
        {
            //Update the client state in the user interface
            clientStateTextBox.Text = currentState.ToString();

            if (currentState == ClientState.SignedIn)
            {
                //Listen for events of changes of the contact's information
                lyncClient.Self.Contact.ContactInformationChanged +=
                    new EventHandler<ContactInformationChangedEventArgs>(SelfContact_ContactInformationChanged);

                //Get the contact's information from Lync and update with it the corresponding elements of the user interface.
                SetName();
                SetAvailability();
                SetPersonalNote();
                SetContactPhoto();

                //Update the SignInOut button content
                signInOutButton.Content = "Sign Out";

                //Enable elements in the user interface
                personalNoteTextBox.IsEnabled = true;
                availabilityComboBox.IsEnabled = /*true;*/ false; // leave the availability box disabled since it doesn't do anything right now
                setNoteButton.IsEnabled = true;
            }
            else
            {
                //Update the SignInOut button content
                signInOutButton.Content = "Sign In";

                //Disable elements in the user interface
                personalNoteTextBox.IsEnabled = false;
                availabilityComboBox.IsEnabled = false;
                setNoteButton.IsEnabled = false;

                //Change the color of the border containing the contact's photo to match the contact's offline status
                availabilityBorder.Background = Brushes.LightSlateGray;
            }
        }

        enum MyColor
        {
            Yellow,
            Red,
            DarkRed,
            LimeGreen,
            LightSlateGray,
            Magenta,
            Blue,
            LightshowMode
        }

        /// <summary>
        /// Gets the contact's current availability value from Lync and updates the corresponding elements in the user interface
        /// </summary>
        private void SetAvailability()
        {
            //Get the current availability value from Lync
            ContactAvailability currentAvailability = 0;
            try
            {
                currentAvailability = (ContactAvailability)
                                                          lyncClient.Self.Contact.GetContactInformation(ContactInformationType.Availability);
            }
            catch (LyncClientException e)
            {
                Console.WriteLine(e);
            }
            catch (SystemException systemException)
            {
                if (IsLyncException(systemException))
                {
                    // Log the exception thrown by the Lync Model API.
                    Console.WriteLine("Error: " + systemException);
                }
                else
                {
                    // Rethrow the SystemException which did not come from the Lync Model API.
                    throw;
                }
            }


            if (currentAvailability != 0)
            {
                //Update the availability ComboBox with the contact's current availability.
                availabilityComboBox.SelectedValue = currentAvailability;

                //Choose a color to match the contact's current availability and update the border area containing the contact's photo
                Brush availabilityColor;
                MyColor myColor;
                switch (currentAvailability)
                {
                    case ContactAvailability.Away:
                        availabilityColor = Brushes.Yellow;
                        myColor = MyColor.Yellow;
                        break;
                    case ContactAvailability.Busy:
                        availabilityColor = Brushes.Red;
                        myColor = MyColor.Red;
                        break;
                    case ContactAvailability.BusyIdle:
                        availabilityColor = Brushes.Red;
                        myColor = MyColor.Red;
                        break;
                    case ContactAvailability.DoNotDisturb:
                        availabilityColor = Brushes.DarkRed;
                        myColor = MyColor.DarkRed;
                        break;
                    case ContactAvailability.Free:
                        availabilityColor = Brushes.LimeGreen;
                        myColor = MyColor.LimeGreen;
                        break;
                    case ContactAvailability.FreeIdle:
                        availabilityColor = Brushes.LimeGreen;
                        myColor = MyColor.LimeGreen;
                        break;
                    case ContactAvailability.Offline:
                        availabilityColor = Brushes.LightSlateGray;
                        myColor = MyColor.LightSlateGray;
                        break;
                    case ContactAvailability.TemporarilyAway:
                        availabilityColor = Brushes.Yellow;
                        myColor = MyColor.Yellow;
                        break;
                    default:
                        availabilityColor = Brushes.LightSlateGray;
                        myColor = MyColor.LightSlateGray;
                        break;
                }
                availabilityBorder.Background = availabilityColor;

                ApplyColor(myColor);
            }
        }

        private void ApplyColor(MyColor color)
        {
            if (serialPort == null || (isLightshowModeActive && color != MyColor.LightshowMode))
            {
                return;
            }

            try
            {
                switch (color)
                {
                    case MyColor.Magenta:
                        {
                            byte[] dat = System.Text.Encoding.ASCII.GetBytes("m");
                            serialPort.Write(dat, 0, 1);
                        }
                        break;

                    case MyColor.Blue:
                        {
                            byte[] dat = System.Text.Encoding.ASCII.GetBytes("b");
                            serialPort.Write(dat, 0, 1);
                        }
                        break;

                    case MyColor.Yellow:
                        {
                            byte[] dat = System.Text.Encoding.ASCII.GetBytes("y");
                            serialPort.Write(dat, 0, 1);
                        }
                        break;

                    case MyColor.Red:
                        {
                            byte[] dat = System.Text.Encoding.ASCII.GetBytes("r");
                            serialPort.Write(dat, 0, 1);
                        }
                        break;

                    case MyColor.DarkRed:
                        {
                            byte[] dat = System.Text.Encoding.ASCII.GetBytes("r");
                            serialPort.Write(dat, 0, 1);
                        }
                        break;

                    case MyColor.LimeGreen:
                        {
                            byte[] dat = System.Text.Encoding.ASCII.GetBytes("g");
                            serialPort.Write(dat, 0, 1);
                        }
                        break;

                    case MyColor.LightSlateGray:
                        {
                            byte[] dat = System.Text.Encoding.ASCII.GetBytes("p");
                            serialPort.Write(dat, 0, 1);
                        }
                        break;

                    default:
                        {
                            byte[] dat = System.Text.Encoding.ASCII.GetBytes("f");
                            serialPort.Write(dat, 0, 1);
                        }
                        break;
                }
            }
            catch (IOException e)
            {
                Console.WriteLine(e);
            }
        }

        /// <summary>
        /// Gets the contact's name from Lync and updates the corresponding element in the user interface
        /// </summary>
        private void SetName()
        {
            string text = string.Empty;
            try
            {
                text = lyncClient.Self.Contact.GetContactInformation(ContactInformationType.DisplayName)
                              as string;
            }
            catch (LyncClientException e)
            {
                Console.WriteLine(e);
            }
            catch (SystemException systemException)
            {
                if (IsLyncException(systemException))
                {
                    // Log the exception thrown by the Lync Model API.
                    Console.WriteLine("Error: " + systemException);
                }
                else
                {
                    // Rethrow the SystemException which did not come from the Lync Model API.
                    throw;
                }
            }

            nameTextBlock.Text = text;
        }

        /// <summary>
        /// Gets the contact's personal note from Lync and updates the corresponding element in the user interface
        /// </summary>
        private void SetPersonalNote()
        {
            string text = string.Empty;
            try
            {
                text = lyncClient.Self.Contact.GetContactInformation(ContactInformationType.PersonalNote)
                              as string;
            }
            catch (LyncClientException e)
            {
                Console.WriteLine(e);
            }
            catch (SystemException systemException)
            {
                if (IsLyncException(systemException))
                {
                    // Log the exception thrown by the Lync Model API.
                    Console.WriteLine("Error: " + systemException);
                }
                else
                {
                    // Rethrow the SystemException which did not come from the Lync Model API.
                    throw;
                }
            }

            personalNoteTextBox.Text = text;
        }

        /// <summary>
        /// Gets the contact's photo from Lync and updates the corresponding element in the user interface
        /// </summary>
        private void SetContactPhoto()
        {
            try
            {
                using (Stream photoStream =
                    lyncClient.Self.Contact.GetContactInformation(ContactInformationType.Photo) as Stream)
                {
                    if (photoStream != null)
                    {
                        BitmapImage bm = new BitmapImage();
                        bm.BeginInit();
                        bm.StreamSource = photoStream;
                        bm.EndInit();
                        photoImage.Source = bm;
                    }
                }
            }
            catch (LyncClientException e)
            {
                Console.WriteLine(e);
            }
            catch (SystemException systemException)
            {
                if (IsLyncException(systemException))
                {
                    // Log the exception thrown by the Lync Model API.
                    Console.WriteLine("Error: " + systemException);
                }
                else
                {
                    // Rethrow the SystemException which did not come from the Lync Model API.
                    throw;
                }
            }
        }

        /// <summary>
        /// Identify if a particular SystemException is one of the exceptions which may be thrown
        /// by the Lync Model API.
        /// </summary>
        /// <param name="ex"></param>
        /// <returns></returns>
        private bool IsLyncException(SystemException ex)
        {
            return
                ex is NotImplementedException ||
                ex is ArgumentException ||
                ex is NullReferenceException ||
                ex is NotSupportedException ||
                ex is ArgumentOutOfRangeException ||
                ex is IndexOutOfRangeException ||
                ex is InvalidOperationException ||
                ex is TypeLoadException ||
                ex is TypeInitializationException ||
                ex is InvalidComObjectException ||
                ex is InvalidCastException;
        }
    }
}

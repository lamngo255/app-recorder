using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Forms;
using System.Windows.Input;

namespace MouseKeyPlayback
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        #region Private Properties
        private KeyboardHook keyboardHook = new KeyboardHook();
        private MouseHook mouseHook = new MouseHook();
        private int count = 0;
        private bool isHooked = false;
        private List<Record> recordList;
        #endregion 


        public MainWindow()
        {
            InitializeComponent();
            recordList = new List<Record>();
            ((INotifyCollectionChanged)listView.Items).CollectionChanged += ListView_CollectionChanged;
        }

        private void ListView_CollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            // Scroll to the last item
            if (e.Action == NotifyCollectionChangedAction.Add)
            {
                listView.ScrollIntoView(e.NewItems[0]);
            }
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            //// Setup keyboard hook
            //keyboardHook.OnKeyboardEvent += KeyboardHook_OnKeyboardEvent;
            //keyboardHook.Install();

            //// Setup mouse hook
            //mouseHook.OnMouseEvent += MouseHook_OnMouseEvent;
            //mouseHook.OnMouseMove += MouseHook_OnMouseMove;
            //mouseHook.OnMouseWheelEvent += MouseHook_OnMouseWheelEvent;
            //mouseHook.Install();
            keyboardHook.OnKeyboardEvent += KeyboardHook_OnKeyboardEvent;

            mouseHook.OnMouseEvent += MouseHook_OnMouseEvent;
            mouseHook.OnMouseMove += MouseHook_OnMouseMove;
            mouseHook.OnMouseWheelEvent += MouseHook_OnMouseWheelEvent;
        }

        private void Window_Closed(object sender, EventArgs e)
        {
            keyboardHook.Uninstall();
            mouseHook.Uninstall();
        }

        #region Mouse events
        private void ProcessMouseEvent(MouseHook.MouseEvents mAction, int mValue)
        {
            CursorPoint mPoint = GetCurrentMousePosition();
            MouseEvent mEvent = new MouseEvent
            {
                Location = mPoint,
                Action = mAction,
                Value = mValue
            };

            LogMouseEvents(mEvent);
        }

        private bool MouseHook_OnMouseWheelEvent(int wheelValue)
        {
            ProcessMouseEvent((MouseHook.MouseEvents)wheelValue, 120);
            return false;
        }

        private bool MouseHook_OnMouseEvent(int mouseEvent)
        {
            ProcessMouseEvent((MouseHook.MouseEvents)mouseEvent, 0);
            return false;
        }


        private bool MouseHook_OnMouseMove(int x, int y)
        {
            ProcessMouseEvent(MouseHook.MouseEvents.MouseMove, 0);
            return false;
        }
        #endregion

        #region Keyboard events
        private bool KeyboardHook_OnKeyboardEvent(uint key, BaseHook.KeyState keyState)
        {
            KeyboardEvent kEvent = new KeyboardEvent {
                Key = (Keys)key,
                Action = (keyState == BaseHook.KeyState.Keydown) ? Constants.KEY_DOWN : Constants.KEY_UP
            };
            LogKeyboardEvents(kEvent);
            return false;
        }
        #endregion

        #region Record/Stop
        private void BtnRecord_Click(object sender, RoutedEventArgs e)
        {
            if (isHooked)
                return;
            if (listView.Items.Count > 0)
            {
                //MessageBoxResult result = System.Windows.MessageBox.Show("Do you want to record again?",
                //                          "Confirmation",
                //                          MessageBoxButton.YesNo,
                //                          MessageBoxImage.Question);
                //switch (result)
                //{
                //    case MessageBoxResult.Yes:
                //        listView.Items.Clear();
                //        recordList = new List<Record>();
                //        count = 0;
                //        break;
                //    default:
                //        return;
                //}

                listView.Items.Clear();
                recordList = new List<Record>();
                count = 0;
            }

            keyboardHook.Install();
            mouseHook.Install();
            isHooked = true;

            LaunchApp();
        }

        private void BtnStop_Click(object sender, RoutedEventArgs e)
        {
            keyboardHook.Uninstall();
            mouseHook.Uninstall();
            isHooked = false;
        }
        #endregion

        #region Helper + Logging methods

        private void LaunchApp()
        {
            // An app is supposed to launch
            if (appPath.IsEnabled == false)
            {
                System.Diagnostics.Process.Start(appPath.Text);
            }
        }
        private void TrackAutomationElement(Record item)
        {
            if (item.Type == Constants.MOUSE
                && item.EventMouse.Action == MouseHook.MouseEvents.LeftUp)
            {
                var windowTitle = Win32Utils.GetActiveWindowTitle();
                var position = Control.MousePosition;

                Point coordinates = new Point(position.X, position.Y);
                inspectBox.Text = string.Format("Title: {0}", windowTitle);

                try
                {
                    AutomationElement targetApp = AutomationElement.FromPoint(coordinates);

                    inspectBox.Text += "\n";
                    inspectBox.Text += string.Format("" +
                        "Name: {0}\n" +
                        "Automation Id: {1}\n" +
                        "Text: {2}\n" +
                        "Control Type: {3}",
                        targetApp.Current.Name,
                        targetApp.Current.AutomationId,
                        GetText(targetApp),
                        targetApp.Current.ControlType);
                }
                catch (Exception)
                {
                    Console.WriteLine("Invalid UI element");
                }
            }
        }

        private string GetText(AutomationElement element)
        {
            object patternObj;
            if (element.TryGetCurrentPattern(ValuePattern.Pattern, out patternObj))
            {
                var valuePattern = (ValuePattern)patternObj;
                return valuePattern.Current.Value;
            }
            else if (element.TryGetCurrentPattern(TextPattern.Pattern, out patternObj))
            {
                var textPattern = (TextPattern)patternObj;
                return textPattern.DocumentRange.GetText(-1).TrimEnd('\r'); // often there is an extra '\r' hanging off the end.
            }
            else
            {
                return element.Current.Name;
            }
        }

        private CursorPoint GetCurrentMousePosition()
        {
            var position = Control.MousePosition;
            return new CursorPoint(position.X, position.Y);
        }

        private void LogMouseEvents(MouseEvent mEvent)
        {
            count++;
            Record item = new Record
            {
                Id = count,
                EventMouse = mEvent,
                Type = Constants.MOUSE,
                Content = String.Format("{0} was triggered at ({1}, {2})", mEvent.Action, mEvent.Location.X, mEvent.Location.Y)
            };

            AddRecordItem(item);
        }

        private void LogKeyboardEvents(KeyboardEvent kEvent)
        {
            count++;
            Record item = new Record
            {
                Id = count,
                Type = Constants.KEYBOARD,
                EventKey = kEvent,
                Content = String.Format("{0} was {1}", kEvent.Key.ToString(),
                    (kEvent.Action == Constants.KEY_DOWN) ? "pressed" : "released")
            };

            AddRecordItem(item);
        }

        private void AddRecordItem(Record item)
        {
            TrackAutomationElement(item);

            AddToListView(item);
            //this.listView.Items.Add(item);
            this.recordList.Add(item);
            countRecord.Content = String.Format("{0} records", count.ToString());
        }

        private void AddToListView(Record item)
        {
            // Check if two last records are similar
            if (listView.Items.Count > 0)
            {
                var lastItem = (Record)listView.Items[listView.Items.Count - 1];
                if (lastItem.Type == item.Type)
                {
                    switch (item.Type)
                    {
                        case Constants.MOUSE:
                            var lastAction = lastItem.EventMouse.Action;
                            if (lastAction == MouseHook.MouseEvents.MouseMove
                                && item.EventMouse.Action == lastAction)
                                this.listView.Items.RemoveAt(this.listView.Items.Count - 1);
                            break;
                        case Constants.KEYBOARD:
                            break;
                    }
                }
            }

            // satisfy every condition
            this.listView.Items.Add(item);
        }
        #endregion

        #region Playback
        private void BtnPlayback_Click(object sender, RoutedEventArgs e)
        {
            if (isHooked)
                return;
            LaunchApp();
            foreach (var record in recordList)
            {
                switch (record.Type)
                {
                    case Constants.MOUSE:
                        PlaybackMouse(record);
                        break;
                    case Constants.KEYBOARD:
                        PlaybackKeyboard(record);
                        break;
                    default:
                        break;
                }
                Thread.Sleep(5);
            }
        }

        private void PlaybackMouse(Record record)
        {
            CursorPoint newPos = record.EventMouse.Location;
            MouseHook.MouseEvents mEvent = record.EventMouse.Action;
            MouseUtils.PerformMouseEvent(mEvent, newPos);
        }

        private void PlaybackKeyboard(Record record)
        {
            Keys key = record.EventKey.Key;
            string action = record.EventKey.Action;

            KeyboardUtils.PerformKeyEvent(key, action);
        }
        #endregion

        private void BtnOpenFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openDialog = new OpenFileDialog();
            openDialog.FileName = "";
            openDialog.Title = "Application";
            openDialog.Filter = "All applications|*.exe";
            openDialog.ShowDialog();

            var appName = openDialog.FileName.ToString();
            if (!String.IsNullOrEmpty(appName))
            {
                appPath.Text = appName;
                appPath.IsEnabled = false;
            } 
        }

        private void Label_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            appPath.Clear();
            appPath.IsEnabled = true;
        }
    }
}

using System;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using MahApps.Metro;
using MahApps.Metro.Controls;
using MahApps.Metro.Controls.Dialogs;
using MetroDemo.Events;
using MetroDemo.ExampleViews;
using MetroDemo.ExampleWindows;
using Microsoft.OfficeProPlus.InstallGen.Presentation.Logging;

namespace MetroDemo
{
    public partial class MainWindow
    {
        private bool _shutdown;
        private int _cacheIndex = -1;

        public MainWindow()
        {
            GlobalObjects.ViewModel = new MainWindowViewModel(DialogCoordinator.Instance)
            {
                ConfigXmlParser = new OfficeInstallGenerator.ConfigXmlParser("<Configuration></Configuration>")
            };

            DataContext = GlobalObjects.ViewModel;

            InitializeComponent();

            ThemeManager.TransitionsEnabled = true;

            MainTabControl.SelectionChanged += MainTabControl_SelectionChanged;

            ProductView.TransitionTab += TransitionTab;

            ProductView.InfoMessage += GenerateViewInfoMessage;
            ProductView.ErrorMessage += GenerateView_ErrorMessage;
        }

        private void MainTabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            e.Handled = false;

            if (GlobalObjects.ViewModel.BlockNavigation)
            {
                MainTabControl.SelectedIndex = _cacheIndex;
                return;
            }

            ThemeManager.TransitionsEnabled = MainTabControl.SelectedIndex != 4;

            if (MainTabControl.SelectedIndex > -1)
            {
                ((TabItem) MainTabControl.Items[MainTabControl.SelectedIndex]).IsEnabled = true;

                if (MainTabControl.SelectedIndex < (MainTabControl.Items.Count - 1))
                {
                    ((TabItem) MainTabControl.Items[MainTabControl.SelectedIndex + 1]).IsEnabled = true;
                }
            }

            if (_cacheIndex != MainTabControl.SelectedIndex)
            {
                if (!GlobalObjects.ViewModel.ResetXml)
                {
                   ProductView.UpdateXml();
                }
                GlobalObjects.ViewModel.ResetXml = false;

                _cacheIndex = MainTabControl.SelectedIndex;
            }
        }

        private void TransitionTab(object sender, Events.TransitionTabEventArgs e)
        {
            try
            {
                var newIndex = Convert.ToInt32(((dynamic)sender).Tag);

                if (e.Direction == TransitionTabDirection.Forward)
                {
                    MainTabControl.SelectedIndex = newIndex + 1;
                }
                else
                {
                    MainTabControl.SelectedIndex = newIndex - 1;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("ERROR:" + ex.Message);
            }
        }

        private async Task ShowMessageDialogAsync(string title, string message)
        {
            await Dispatcher.InvokeAsync(async () =>
            {
                var result = await this.ShowMessageAsync(title, message, 
                    MessageDialogStyle.Affirmative, new MetroDialogSettings()
                {
                    ColorScheme = MetroDialogColorScheme.Theme
                });
            });
        }

        private async Task ShowErrorDialogAsync(string title, string message)
        {
            await Dispatcher.InvokeAsync(async () =>
            {
                var result = await this.ShowMessageAsync(title, message,
                    MessageDialogStyle.Affirmative, new MetroDialogSettings()
                    {
                        ColorScheme = MetroDialogColorScheme.Error
                    });
            });
        }

        #region Events

        private async void GenerateViewInfoMessage(object sender, MessageEventArgs e)
        {
            try
            {
                await ShowMessageDialogAsync(e.Title, e.Message);
            }
            catch (Exception ex)
            {
                ShowErrorDialogAsync("ERROR", ex.Message).ConfigureAwait(false);
            }
        }

        private async void GenerateView_ErrorMessage(object sender, MessageEventArgs e)
        {
            try
            {
                await ShowErrorDialogAsync(e.Title, e.Message);
            }
            catch (Exception ex)
            {
                ShowErrorDialogAsync("ERROR", ex.Message).ConfigureAwait(false);
            }
        }

        #endregion

        #region Other
        public static readonly DependencyProperty ToggleFullScreenProperty =
            DependencyProperty.Register("ToggleFullScreen",
                                        typeof(bool),
                                        typeof(MainWindow),
                                        new PropertyMetadata(default(bool), ToggleFullScreenPropertyChangedCallback));

        private static void ToggleFullScreenPropertyChangedCallback(DependencyObject dependencyObject, DependencyPropertyChangedEventArgs e)
        {
            var metroWindow = (MetroWindow)dependencyObject;
            if (e.OldValue != e.NewValue)
            {
                var fullScreen = (bool)e.NewValue;
                if (fullScreen)
                {
                    metroWindow.UseNoneWindowStyle = true;
                    metroWindow.IgnoreTaskbarOnMaximize = true;
                    metroWindow.WindowState = WindowState.Maximized;
                }
                else
                {
                    metroWindow.UseNoneWindowStyle = false;
                    metroWindow.ShowTitleBar = true; // <-- this must be set to true
                    metroWindow.IgnoreTaskbarOnMaximize = false;
                    metroWindow.WindowState = WindowState.Normal;
                }
            }
        }

        public bool ToggleFullScreen
        {
            get { return (bool)GetValue(ToggleFullScreenProperty); }
            set { SetValue(ToggleFullScreenProperty, value); }
        }

        public static readonly DependencyProperty UseAccentForDialogsProperty =
            DependencyProperty.Register("UseAccentForDialogs",
                                        typeof(bool),
                                        typeof(MainWindow),
                                        new PropertyMetadata(default(bool), ToggleUseAccentForDialogsPropertyChangedCallback));

        private static void ToggleUseAccentForDialogsPropertyChangedCallback(DependencyObject dependencyObject, DependencyPropertyChangedEventArgs e)
        {
            var metroWindow = (MetroWindow)dependencyObject;
            if (e.OldValue != e.NewValue)
            {
                var useAccentForDialogs = (bool)e.NewValue;
                metroWindow.MetroDialogOptions.ColorScheme = useAccentForDialogs ? MetroDialogColorScheme.Accented : MetroDialogColorScheme.Theme;
            }
        }

        public bool UseAccentForDialogs
        {
            get { return (bool)GetValue(UseAccentForDialogsProperty); }
            set { SetValue(UseAccentForDialogsProperty, value); }
        }

        private async void CloseCustomDialog(object sender, RoutedEventArgs e)
        {
            var dialog = (BaseMetroDialog)this.Resources["CustomCloseDialogTest"];

            await this.HideMetroDialogAsync(dialog);
        }

        private async void MetroWindow_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            try
            {
                e.Cancel = !_shutdown;
                if (_shutdown) return;

                var mySettings = new MetroDialogSettings()
                {
                    AffirmativeButtonText = "Quit",
                    NegativeButtonText = "Cancel",
                    AnimateShow = true,
                    AnimateHide = false
                };

                var result = await this.ShowMessageAsync("Quit application?",
                    "Sure you want to quit application?",
                    MessageDialogStyle.AffirmativeAndNegative, mySettings);

                _shutdown = result == MessageDialogResult.Affirmative;

                if (_shutdown)
                {
                    Application.Current.Shutdown();
                }
            }
            catch { }
        }
        #endregion
        
    }
}

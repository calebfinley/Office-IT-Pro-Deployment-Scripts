﻿using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.Remoting.Channels;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using MahApps.Metro.Controls;
using MetroDemo.Events;
using MetroDemo.ExampleWindows;
using MetroDemo.Models;
using Micorosft.OfficeProPlus.ConfigurationXml;
using Micorosft.OfficeProPlus.ConfigurationXml.Model;
using Microsoft.OfficeProPlus.Downloader;
using Microsoft.OfficeProPlus.Downloader.Model;
using Microsoft.OfficeProPlus.InstallGen.Presentation.Logging;
using Microsoft.OfficeProPlus.InstallGen.Presentation.Models;
using Microsoft.OfficeProPlus.InstallGenerator.Models;
using OfficeInstallGenerator.Model;
using File = System.IO.File;
using MessageBox = System.Windows.MessageBox;
using UserControl = System.Windows.Controls.UserControl;

namespace MetroDemo.ExampleViews
{
    /// <summary>
    /// Interaction logic for TextExamples.xaml
    /// </summary>
    public partial class ProductView : UserControl
    {
        private LanguagesDialog languagesDialog = null;
        private CancellationTokenSource _tokenSource = new CancellationTokenSource();

        public event TransitionTabEventHandler TransitionTab;
        public event MessageEventHandler InfoMessage;
        public event MessageEventHandler ErrorMessage;

        private Task _downloadTask = null;
        private int _cachedIndex = 0;
        
        public ProductView()
        {
            InitializeComponent();
        }

        private void ProductView_Loaded(object sender, RoutedEventArgs e)             
        {
            try
            {
                if (MainTabControl == null) return;
                MainTabControl.SelectedIndex = 0;

                if (GlobalObjects.ViewModel == null) return;
                LanguageList.ItemsSource = GlobalObjects.ViewModel.GetLanguages(null);
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        private void LogAnaylytics(string path, string pageName)
        {
            try
            {
                //GoogleAnalytics.Log(path, pageName);
            }
            catch { }
        }

        private async Task DownloadOfficeFiles()
        {
            try
            {
                SetTabStatus(false);
                GlobalObjects.ViewModel.BlockNavigation = true;
                _tokenSource = new CancellationTokenSource();

                UpdateXml();

                ProductUpdateSource.IsReadOnly = true;
                UpdatePath.IsEnabled = false;

                DownloadProgressBar.Maximum = 100;
                DownloadPercent.Content = "";

                var configXml = GlobalObjects.ViewModel.ConfigXmlParser.ConfigurationXml;

                string branch = null;
                if (configXml.Add.Branch.HasValue)
                {
                    branch = configXml.Add.Branch.Value.ToString();
                }

                var proPlusDownloader = new ProPlusDownloader();
                proPlusDownloader.DownloadFileProgress += async (senderfp, progress) =>
                {
                    var percent = progress.PercentageComplete;
                    if (percent > 0)
                    {
                        Dispatcher.Invoke(() =>
                        {
                            DownloadPercent.Content = percent + "%";
                            DownloadProgressBar.Value = Convert.ToInt32(Math.Round(percent, 0));
                        });
                    }
                };
                proPlusDownloader.VersionDetected += (sender, version) =>
                {
                    if (branch == null) return;
                    var modelBranch = GlobalObjects.ViewModel.Branches.FirstOrDefault(b => b.Branch.ToString().ToLower() == branch.ToLower());
                    if (modelBranch == null) return;
                    if (modelBranch.Versions.Any(v => v.Version == version.Version)) return;
                    modelBranch.Versions.Insert(0, new Build() { Version = version.Version });
                    modelBranch.CurrentVersion = version.Version;

                    ProductVersion.ItemsSource = modelBranch.Versions;
                    ProductVersion.SetValue(TextBoxHelper.WatermarkProperty, modelBranch.CurrentVersion);
                };

                var buildPath = ProductUpdateSource.Text.Trim();
                if (string.IsNullOrEmpty(buildPath)) return;

                var languages =
                    (from product in configXml.Add.Products
                     from language in product.Languages
                     select language.ID.ToLower()).Distinct().ToList();

                var officeEdition = OfficeEdition.Office32Bit;
                if (configXml.Add.OfficeClientEdition == OfficeClientEdition.Office64Bit)
                {
                    officeEdition = OfficeEdition.Office64Bit;
                }

                buildPath = GlobalObjects.SetBranchFolderPath(branch, buildPath);
                Directory.CreateDirectory(buildPath);

                ProductUpdateSource.Text = buildPath;

                await proPlusDownloader.DownloadBranch(new DownloadBranchProperties()
                {
                    BranchName = branch,
                    OfficeEdition = officeEdition,
                    TargetDirectory = buildPath,
                    Languages = languages
                }, _tokenSource.Token);

                MessageBox.Show("Download Complete");

                LogAnaylytics("/ProductView", "Download." + branch);
            }
            finally
            {
                SetTabStatus(true);
                GlobalObjects.ViewModel.BlockNavigation = false;
                ProductUpdateSource.IsReadOnly = false;
                UpdatePath.IsEnabled = true;
                DownloadProgressBar.Value = 0;
                DownloadPercent.Content = "";

                DownloadButton.Content = "Download";
                _tokenSource = new CancellationTokenSource();
            }
        }


        private void LanguageChange()
        {
            var languages = GlobalObjects.ViewModel.GetLanguages(GetSelectedProduct());

            LanguageList.ItemsSource = null;
            LanguageList.ItemsSource = languages;
        }

        private void RemoveSelectedLanguage()
        {
            var selectProductId = GetSelectedProduct();

            var currentItems = (List<Language>)LanguageList.ItemsSource ?? new List<Language>();
            foreach (Language language in LanguageList.SelectedItems)
            {
                if (currentItems.Contains(language))
                {
                    currentItems.Remove(language);
                }



                GlobalObjects.ViewModel.RemoveLanguage(selectProductId, language.Id);
            }
            LanguageList.ItemsSource = null;
            LanguageList.ItemsSource = GlobalObjects.ViewModel.GetLanguages(selectProductId);
        }

        private void ChangePrimaryLanguage()
        {
            var currentItems = (List<Language>)LanguageList.ItemsSource ?? new List<Language>();
            if (currentItems.Count <= 0) return;
            if (LanguageList.SelectedItems.Count != 1) return;

            var selectedLanguage = LanguageList.SelectedItems.Cast<Language>().FirstOrDefault();

            var selectProductId = GetSelectedProduct();

            GlobalObjects.ViewModel.ChangePrimaryLanguage(selectProductId, selectedLanguage);

            LanguageList.ItemsSource = null;
            LanguageList.ItemsSource = GlobalObjects.ViewModel.GetLanguages(selectProductId);
        }

        public void Reset()
        {
            ProductEditionAll.IsChecked = true;
            ProductEdition32Bit.IsChecked = false;
            ProductEdition64Bit.IsChecked = false;
            ProductVersion.Text = "";
            ProductUpdateSource.Text = "";

            GlobalObjects.ViewModel.ClearLanguages();
        }

        public void LoadXml()
        {
            var languages = new List<Language>
            {
                GlobalObjects.ViewModel.DefaultLanguage
            };

            Reset();

            var configXml = GlobalObjects.ViewModel.ConfigXmlParser.ConfigurationXml;
            if (configXml.Add != null)
            {
                if (configXml.Add.OfficeClientEdition == OfficeClientEdition.Office32Bit)
                {
                    ProductEdition32Bit.IsChecked = true;
                    ProductEdition64Bit.IsChecked = false;
                }
                if (configXml.Add.OfficeClientEdition == OfficeClientEdition.Office64Bit)
                {
                    ProductEdition32Bit.IsChecked = false;
                    ProductEdition64Bit.IsChecked = true;
                }

                ProductVersion.Text = configXml.Add.Version != null ? configXml.Add.Version.ToString() : "";
                ProductUpdateSource.Text = configXml.Add.SourcePath != null ? configXml.Add.SourcePath.ToString() : "";

                if (configXml.Add.Products != null && configXml.Add.Products.Count > 0)
                {
                    LanguageList.ItemsSource = null;

                    GlobalObjects.ViewModel.ClearLanguages();

                    var n = 0;
                    foreach (var product in configXml.Add.Products)
                    {
                        var index = 0;

                        if (product.Languages != null)
                        {
                            if (n == 0) languages.Clear();

                            var useSameLangs = configXml.Add.IsLanguagesSameForAllProducts();

                            var order = 1;
                            foreach (var language in product.Languages)
                            {
                                var languageLookup = GlobalObjects.ViewModel.Languages.FirstOrDefault(
                                                        l => l.Id.ToLower() == language.ID.ToLower());
                                if (languageLookup == null) continue;
                                string productId = null;

                                if (!useSameLangs)
                                {
                                    productId = product.ID;
                                }

                                var newLanguage = new Language()
                                {
                                    Id = languageLookup.Id,
                                    Name = languageLookup.Name,
                                    Order = order,
                                    ProductId = productId
                                };

                                GlobalObjects.ViewModel.AddLanguage(productId, newLanguage);

                                if (n == 0) languages.Add(newLanguage);
                                order++;
                            }

                            n++;
                        }
                    }


                }
            }
            else
            {
                ProductEdition32Bit.IsChecked = true;
                ProductEdition64Bit.IsChecked = false;
                ProductVersion.Text = "";
            }

            var distictList = languages.Distinct().ToList();
            LanguageList.ItemsSource = FormatLanguage(distictList);

            LanguageChange();
        }

        public void UpdateXml()
        {
            var configXml = GlobalObjects.ViewModel.ConfigXmlParser.ConfigurationXml;
            if (configXml.Add == null)
            {
                configXml.Add = new ODTAdd();
            }

            if (ProductEdition32Bit.IsChecked.HasValue && ProductEdition32Bit.IsChecked.Value)
            {
                configXml.Add.OfficeClientEdition = OfficeClientEdition.Office32Bit;
            }

            if (ProductEdition64Bit.IsChecked.HasValue && ProductEdition64Bit.IsChecked.Value)
            {
                configXml.Add.OfficeClientEdition = OfficeClientEdition.Office64Bit;
            }

            if (configXml.Add.Products == null)
            {
                configXml.Add.Products = new List<ODTProduct>();   
            }

            var versionText = "";
            if (ProductVersion.SelectedIndex > -1)
            {
                var version = (Build) ProductVersion.SelectedValue;
                versionText = version.Version;
            }
            else
            {
                versionText = ProductVersion.Text;
            }

            try
            {
                if (!string.IsNullOrEmpty(versionText))
                {
                    Version productVersion = null;
                    Version.TryParse(versionText, out productVersion);
                    configXml.Add.Version = productVersion;
                }
                else
                {
                    configXml.Add.Version = null;
                }
            }
            catch { }

            configXml.Add.SourcePath = ProductUpdateSource.Text.Length > 0 ? ProductUpdateSource.Text : null;
        }


        public async Task UpdateVersions()
        {
        }

        private IEnumerable<Language> FormatLanguage(List<Language> languages)
        {
            if (languages == null) return new List<Language>();
            foreach (var language in languages)
            {
                language.Name = Regex.Replace(language.Name, @"\s\(Primary\)", "", RegexOptions.IgnoreCase);
            }
            if (languages.Any())
            {
                languages.FirstOrDefault().Name += " (Primary)";
            }
            return languages.ToList();
        }


        private string GetSelectedProduct()
        {
            string selectedProductId = null;
            return selectedProductId;
        }

        private async Task GetBranchVersion(OfficeBranch branch, OfficeEdition officeEdition)
        {
            try
            {
                if (branch.Updated) return;
                var ppDownload = new ProPlusDownloader();
                var latestVersion = await ppDownload.GetLatestVersionAsync(branch.Branch.ToString(), officeEdition);

                var modelBranch = GlobalObjects.ViewModel.Branches.FirstOrDefault(b =>
                    b.Branch.ToString().ToLower() == branch.Branch.ToString().ToLower());
                if (modelBranch == null) return;
                if (modelBranch.Versions.Any(v => v.Version == latestVersion)) return;
                modelBranch.Versions.Insert(0, new Build() { Version = latestVersion });
                modelBranch.CurrentVersion = latestVersion;

                ProductVersion.ItemsSource = modelBranch.Versions;
                ProductVersion.SetValue(TextBoxHelper.WatermarkProperty, modelBranch.CurrentVersion);

                modelBranch.Updated = true;
            }
            catch (Exception)
            {

            }
        }

        private bool TransitionProductTabs(TransitionTabDirection direction)
        {
            if (direction == TransitionTabDirection.Forward)
            {
                if (MainTabControl.SelectedIndex < MainTabControl.Items.Count - 1)
                {
                    MainTabControl.SelectedIndex++;
                }
                else
                {
                    return true;
                }
            }
            else
            {
                if (MainTabControl.SelectedIndex > 0)
                {
                    MainTabControl.SelectedIndex--;
                }
                else
                {
                    return true;
                }
            }

            return false;
        }

        private void LogErrorMessage(Exception ex)
        {
            ex.LogException(false);
            if (ErrorMessage != null)
            {
                ErrorMessage(this, new MessageEventArgs()
                {
                    Title = "Error",
                    Message = ex.Message
                });
            }
        }

        private void SetTabStatus(bool enabled)
        {
            Dispatcher.Invoke(() =>
            {
                ProductTab.IsEnabled = enabled;
                OptionalTab.IsEnabled = enabled;
            });
        }

        #region "Events"

        private async void DownloadButton_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_tokenSource != null)
                {
                    if (_tokenSource.IsCancellationRequested)
                    {
                        GlobalObjects.ViewModel.BlockNavigation = false;
                        SetTabStatus(true);
                        return;
                    }
                    if (_downloadTask.IsActive())
                    {
                        GlobalObjects.ViewModel.BlockNavigation = false;
                        SetTabStatus(true);
                        _tokenSource.Cancel();
                        return;
                    }
                }

                DownloadButton.Content = "Stop";

                _downloadTask = DownloadOfficeFiles();
                await _downloadTask;
            }
            catch (Exception ex)
            {
                if (ex.Message.ToLower().Contains("aborted") ||
                    ex.Message.ToLower().Contains("canceled"))
                {
                    GlobalObjects.ViewModel.BlockNavigation = false;
                    SetTabStatus(true);
                }
                else
                {
                    LogErrorMessage(ex);
                }
            }
        }

        private async void OpenFolderButton_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                var folderPath = ProductUpdateSource.Text.Trim();
                if (string.IsNullOrEmpty(folderPath)) return;

                if (await GlobalObjects.DirectoryExists(folderPath))
                {
                    Process.Start("explorer", folderPath);
                }
                else
                {
                    MessageBox.Show("Directory path does not exist.");
                }
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }


        private async void BuildFilePath_OnTextChanged(object sender, TextChangedEventArgs e)
        {
            try
            {
                var enabled = false;
                var openFolderEnabled = false;
                if (ProductUpdateSource.Text.Trim().Length > 0)
                {
                    var match = Regex.Match(ProductUpdateSource.Text, @"^\w:\\|\\\\.*\\..*");
                    if (match.Success)
                    {
                        enabled = true;
                        var folderExists = await GlobalObjects.DirectoryExists(ProductUpdateSource.Text);
                        if (!folderExists)
                        {
                            folderExists = await GlobalObjects.DirectoryExists(ProductUpdateSource.Text);
                        }

                        openFolderEnabled = folderExists;  
                    }
                }

                OpenFolderButton.IsEnabled = openFolderEnabled;
                DownloadButton.IsEnabled = enabled;
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }
        


        private void UpdatePath_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                var dlg1 = new Ionic.Utils.FolderBrowserDialogEx
                {
                    Description = "Select a folder:",
                    ShowNewFolderButton = true,
                    ShowEditBox = true,
                    SelectedPath = ProductUpdateSource.Text,
                    ShowFullPathInEditBox = true,
                    RootFolder = System.Environment.SpecialFolder.MyComputer
                };
                //dlg1.NewStyle = false;

                // Show the FolderBrowserDialog.
                var result = dlg1.ShowDialog();
                if (result == DialogResult.OK)
                {
                    ProductUpdateSource.Text = dlg1.SelectedPath;
                }
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        private void LanguageUnique_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                LanguageChange();
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }


        private void MainTabControl_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                if (GlobalObjects.ViewModel.BlockNavigation)
                {
                    MainTabControl.SelectedIndex = _cachedIndex;
                    return;
                }

                _cachedIndex = MainTabControl.SelectedIndex;
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        private void ToggleSwitch_OnIsCheckedChanged(object sender, EventArgs e)
        {
            try
            {
                var toggleSwitch = (ToggleSwitch) sender;
                if (toggleSwitch != null)
                {
                    var context = (ExcludeProduct) toggleSwitch.DataContext;
                    if (context != null)
                    {
                        if (toggleSwitch.IsChecked.HasValue)
                        {
                            context.Included = toggleSwitch.IsChecked.Value;

                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        private void AddLanguageButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                LaunchLanguageDialog();
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        private void RemoveLanguageButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                RemoveSelectedLanguage();
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        private void NextButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                UpdateXml();

                if (TransitionProductTabs(TransitionTabDirection.Forward))
                {
                    this.TransitionTab(this, new TransitionTabEventArgs()
                    {
                        Direction = TransitionTabDirection.Forward,
                        Index = 1
                    });
                }
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        private void PreviousButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                UpdateXml();

                if (TransitionProductTabs(TransitionTabDirection.Back))
                {
                    this.TransitionTab(this, new TransitionTabEventArgs()
                    {
                        Direction = TransitionTabDirection.Back,
                        Index = 1
                    });
                }
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }
        
        public BranchChangedEventHandler BranchChanged { get; set; }

        #endregion

        #region "Info"

        private void ProductInfo_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                var sourceName = ((dynamic) sender).Name;
                LaunchInformationDialog(sourceName);
            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        private InformationDialog informationDialog = null;

        private void LaunchInformationDialog(string sourceName)
        {
            try
            {
                if (informationDialog == null)
                {

                    informationDialog = new InformationDialog
                    {
                        Height = 500,
                        Width = 400
                    };
                    informationDialog.Closed += (o, args) =>
                    {
                        informationDialog = null;
                    };
                    informationDialog.Closing += (o, args) =>
                    {

                    };
                }
                
                informationDialog.Height = 500;
                informationDialog.Width = 400;

                var filePath = AppDomain.CurrentDomain.BaseDirectory + @"HelpFiles\" + sourceName + ".html";
                var helpFile = File.ReadAllText(filePath);

                informationDialog.HelpInfo.NavigateToString(helpFile);
                informationDialog.Launch();

            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }

        private void LaunchLanguageDialog()
        {
            try
            {
                if (languagesDialog == null)
                {
                    var currentItems1 = (List<Language>)LanguageList.ItemsSource ?? new List<Language>();

                    var languageList = GlobalObjects.ViewModel.Languages.ToList();
                    foreach (var language in currentItems1)
                    {
                        languageList.Remove(language);
                    }

                    languagesDialog = new LanguagesDialog
                    {
                        LanguageSource = languageList
                    };
                    languagesDialog.Closed += (o, args) =>
                    {
                        languagesDialog = null;
                    };
                    languagesDialog.Closing += (o, args) =>
                    {
                        var currentItems2 = (List<Language>)LanguageList.ItemsSource ?? new List<Language>();

                        if (languagesDialog.SelectedItems != null)
                        {
                            if (languagesDialog.SelectedItems.Count > 0)
                            {
                                currentItems2.AddRange(languagesDialog.SelectedItems);
                            }
                        }

                        var selectedLangs = FormatLanguage(currentItems2.Distinct().ToList()).ToList();

                        var selectProductId = GetSelectedProduct();

                        foreach (var languages in selectedLangs)
                        {
                            languages.ProductId = selectProductId;
                        }

                        GlobalObjects.ViewModel.AddLanguages(selectProductId, selectedLangs);

                        LanguageList.ItemsSource = null;
                        LanguageList.ItemsSource = selectedLangs;

                    };
                }
                languagesDialog.Launch();

            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }
        }


        #endregion


    }
}


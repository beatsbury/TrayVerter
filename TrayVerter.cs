using ImageMagick;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Management;
using System.ServiceProcess;
using System.Threading;

namespace TrayVerter
{
    public partial class TrayVerter : ServiceBase
    {


        EventLog ServiceLog;

        FileSystemWatcher watcher;

        MagickImage image;

        public TrayVerter()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            InitializeService();

            InitializeFSW();
        }

        private void InitializeFSW()
        {            
            watcher = new FileSystemWatcher();

            string currentUser = GetExplorerUser();
            ServiceLog = this.EventLog;
            ServiceLog.WriteEntry(currentUser);
            
            string ThemePathString = System.Environment.GetFolderPath(Environment.SpecialFolder.Windows).Replace("WINDOWS", "") + @"Users\" + currentUser + @"\AppData\Roaming\Microsoft\Windows\Themes\";
            ServiceLog.WriteEntry(ThemePathString);

            watcher.Filter = "TranscodedWallpaper";

            watcher.Path = ThemePathString;
            watcher.NotifyFilter = NotifyFilters.Attributes;

            // Begin watching.
            watcher.EnableRaisingEvents = true;

            // Add event handlers.                
            watcher.Changed += new FileSystemEventHandler(OnChanged);

        }

        private void InitializeService()
        {

        }

        private void ChangeSetting(bool isDark)
        {
            RegistryKey regKey = Registry.CurrentUser;
            ServiceLog.WriteEntry(regKey.ToString());
            RegistryKey regSubKey = regKey.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize", true);
            string CurrentValue = regSubKey.GetValue("SystemUsesLightTheme").ToString();
            string NewValue = isDark ? "0" : "1";
            ServiceLog.WriteEntry("Current value: " + CurrentValue);
            ServiceLog.WriteEntry("New value: " + NewValue);
            regSubKey.SetValue("SystemUsesLightTheme", NewValue);
            ServiceLog.WriteEntry("Value after setting: " + regSubKey.GetValue("SystemUsesLightTheme").ToString());
            ServiceLog.WriteEntry("Reg key is: " + regSubKey.ToString());
            regSubKey.SetValue("AppsUseLightTheme", "1");
        }

        private bool IsDark(string filePath)
        {
            string CurrentWallpaperPath = filePath;

            ServiceLog.WriteEntry("IsDark() fired..");

            Thread.Sleep(500);

            try
            {
                using (image = new MagickImage(CurrentWallpaperPath))
                {
                    IEnumerable<string> regScreenSizeArr = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Windows\Shell\Bags\1\Desktop").GetValueNames().Where(item => item.Contains("ItemPos"));
                    string regScreenSize = regScreenSizeArr.ElementAt(0);
                    string regScreenSizeTrimmed = regScreenSize.Remove(regScreenSize.LastIndexOf('x')).Substring(7);
                    int screenWidth = Convert.ToInt32(regScreenSizeTrimmed.Substring(0, regScreenSizeTrimmed.LastIndexOf('x')));
                    int screenHeight = Convert.ToInt32(regScreenSizeTrimmed.Substring(regScreenSizeTrimmed.LastIndexOf('x') + 1));
                    ServiceLog.WriteEntry("Got screen dimensions: " + screenWidth + "X" + screenHeight);

                    int imageWidth = image.Width;
                    int imageHeight = image.Height;
                    ServiceLog.WriteEntry("Got image dimensions: " + imageWidth + "X" + imageHeight);

                    if (imageWidth != screenWidth)
                    {
                        image.Resize(screenWidth, imageHeight * screenWidth / imageWidth);
                    }

                    string geometry = image.Width.ToString() + "x48+0+" + (screenHeight - 48).ToString();
                    ServiceLog.WriteEntry("Image width after resizing: " + image.Width + "X" + image.Height);
                    ServiceLog.WriteEntry("Geometry for cropping: " + geometry);
                    image.Crop(new MagickGeometry(geometry));
                    ServiceLog.WriteEntry("Resulting image dimensions: " + image.Width + "x" + image.Height);
                    image.Clamp();
                    image.Grayscale();

                    IPixelCollection foundPixels = image.GetPixels();
                    int channelsOverall = foundPixels.Channels;

                    ServiceLog.WriteEntry("Channels overall : " + channelsOverall);
                    ColorType foundColor = image.DetermineColorType();
                    ServiceLog.WriteEntry(foundColor.ToString());

                    int greySum = 0;

                    int horThreshold1 = Convert.ToInt32(image.Width * 0.04);
                    int horThreshold2 = Convert.ToInt32(image.Width * 0.9);

                    int pixelsTotal = 0;

                    foreach (var foundPixel in foundPixels)
                    {
                        if (foundPixel.X <= horThreshold1 || foundPixel.X >= horThreshold2)
                        {
                            int greyValue = foundPixel.GetChannel(0);
                            greySum += greyValue;
                            pixelsTotal += 1;
                        }
                    }

                    ServiceLog.WriteEntry("Pixels total : " + pixelsTotal);
                    ServiceLog.WriteEntry("Sum of all values : " + greySum);

                    int averageValue = greySum / pixelsTotal;
                    ServiceLog.WriteEntry("The average is : " + averageValue);

                    return averageValue < 120 ? true : false;
                }
            }
            catch (MagickException exception)
            {
                ServiceLog.WriteEntry(exception.Message);
                //Console.WriteLine();
                Random rnd = new Random();
                int testInt = rnd.Next(0, 255);
                return testInt < 120 ? true : false;
            }
        }

        private void OnChanged(object sender, FileSystemEventArgs e)
        {
            ServiceLog.WriteEntry("OnChanged() fired..");
            bool isItDark = IsDark(e.FullPath);
            ChangeSetting(isItDark);
        }

        private string GetExplorerUser()
        {
            var process = Process.GetProcessesByName("explorer");
            return process.Length > 0
                ? GetUsernameByPid(process[0].Id)
                : "Unknown-User";
        }

        private string GetUsernameByPid(int pid)
        {
            var query = new ObjectQuery("SELECT * from Win32_Process "
                + " WHERE ProcessID = '" + pid + "'");

            var searcher = new ManagementObjectSearcher(query);
            if (searcher.Get().Count == 0)
                return "Unknown-User";

            foreach (ManagementObject obj in searcher.Get())
            {
                var owner = new String[2];
                obj.InvokeMethod("GetOwner", owner);
                return owner[0] ?? "Unknown-User";
            }

            return "Unknown-User";
        }

        protected override void OnStop()
        {
            watcher.EnableRaisingEvents = false;
            watcher.Dispose();
        }
    }
}

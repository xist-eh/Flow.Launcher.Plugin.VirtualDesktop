using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;

using Flow.Launcher.Plugin;


namespace Flow.Launcher.Plugin.VirtualDesktop
{
    public class Main : IPlugin
    {
        private PluginInitContext _context;
        private const string IconPath = "Images/app.png";


        public void Init(PluginInitContext context)
        {
            _context = context;
        }

        public List<Result> Query(Query query)
        {
            var results = new List<Result>();

            // Execute in STA thread to access COM objects
            int desktopCount = 0;
            List<string> desktopNames = null;
            int currentDesktop = 0;

            // Safely get desktop information using STA thread
            ExecuteStaThread(() =>
            {
                try
                {
                    desktopCount = VDManager.GetDesktopCount();
                    desktopNames = VDManager.GetAllDesktopNames();
                    currentDesktop = VDManager.GetCurrentDesktopIndex();
                }
                catch (Exception)
                {
                    // Fallback to default values if COM calls fail
                    desktopCount = 0;
                    desktopNames = new List<string>();
                }
                return true;
            });

            // Create results based on desktop count
            for (int i = 0; i < desktopCount; i++)
            {
                var desktopIndex = i; // Capture for lambda
                string title = (i < desktopNames?.Count) ? desktopNames[i] : $"Virtual Desktop {i + 1}";
                string subtitle = (desktopIndex == currentDesktop) ?
                    $"Current Desktop (Desktop {i + 1})" : $"Switch to Virtual Desktop {i + 1}";

                results.Add(new Result
                {
                    Title = title,
                    SubTitle = subtitle,
                    IcoPath = IconPath,
                    Action = c =>
                    {
                        return ExecuteStaThread(() =>
                        {
                            try
                            {
                                if (desktopIndex != currentDesktop)
                                {
                                    // Only switch if it's not the current desktop
                                    return VDManager.SwitchToDesktop(desktopIndex);
                                }
                                return true;
                            }
                            catch
                            {
                                return false;
                            }
                        });
                    }
                });
            }

            // Add management option
            results.Add(new Result
            {
                Title = "Virtual Desktop Manager",
                SubTitle = "Manage your virtual desktops",
                IcoPath = IconPath,
                Action = c =>
                {
                    return true;
                }
            });
            return results;
        }

        // Helper method to execute code in an STA thread
        static private bool ExecuteStaThread(Func<bool> action)
        {
            bool result = false;
            var thread = new Thread(() =>
            {
                try
                {
                    result = action();
                }
                catch (Exception ex)
                {
                    result = false;
                }
            });

            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join();

            return result;
        }

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();
    }
}
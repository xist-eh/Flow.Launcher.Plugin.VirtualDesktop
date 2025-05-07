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

            // Create array to store desktop count
            int desktopCount = 10;


            // Create results based on desktop count
            for (int i = 0; i < desktopCount; i++)
            {
                var desktopIndex = i; // Capture for lambda
                results.Add(new Result
                {
                    Title = $"Virtual Desktop {i + 1}",
                    SubTitle = $"Switch to Virtual Desktop {i + 1}",
                    IcoPath = IconPath,
                    Action = c =>
                    {
                        return true;
                    }
                });
            }

            results.Add(new Result
            {
                Title = "Virtual Desktop",
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
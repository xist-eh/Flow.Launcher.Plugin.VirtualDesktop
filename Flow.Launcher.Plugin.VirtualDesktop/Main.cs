using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Diagnostics;
using System.IO;

using Flow.Launcher.Plugin;


namespace Flow.Launcher.Plugin.VirtualDesktop
{
    public class Main : IPlugin
    {
        private PluginInitContext _context;
        private const string IconPath = "Images/vd.ico";
        private string vdExeName = "VirtualDesktop11.exe";
        private string vdFullPath;


        public void Init(PluginInitContext context)
        {
            _context = context;
            int buildNumber = GetWindowsBuildNumber();
            if (buildNumber < 2200)
            {
                vdExeName = "VirtualDesktop.exe";
            }
            else if (buildNumber >= 26100)
            {
                vdExeName = "VirtualDesktop11-24H2.exe";
            }

            // Get the full path to the executable
            string pluginDirectory = _context.CurrentPluginMetadata.PluginDirectory;
            vdFullPath = Path.Combine(pluginDirectory, "vdmanager", vdExeName);
        }

        public List<Result> Query(Query query)
        {
            var results = new List<Result>();

            if (query.SearchTerms.Any(s => s.Contains("debug", StringComparison.OrdinalIgnoreCase)))
            {

                results.Add(new Result { Title = "Build Number", SubTitle = GetWindowsBuildNumber().ToString(), IcoPath = IconPath });
                results.Add(new Result { Title = "VD Path", SubTitle = vdFullPath, IcoPath = IconPath });
            }
            string[] desktops = GetAllDesktops();
            int currentDesktopIndex = CallVDManager("/Q /GetCurrentDesktop").ExitCode;
            for (int i = 0; i < desktops.Length; i++)
            {
                if (i == currentDesktopIndex)
                {
                    continue; // Skip the current desktop
                }
                string desktop = desktops[i];
                int index = i;  // Create a local copy of the loop variable

                results.Add(new Result
                {
                    Title = desktop,
                    SubTitle = "Switch to this desktop",
                    IcoPath = IconPath,
                    Action = (e) =>
                    {
                        CallVDManager($"/Switch:{index}");
                        return true;
                    }
                });
            }


            return results;
        }

        private string[] GetAllDesktops()
        {
            return CallVDManager("/Q /List").Output;
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

        private ProcessResult CallVDManager(string args)
        {
            var processInfo = new ProcessStartInfo
            {
                FileName = vdFullPath,
                Arguments = args,
                RedirectStandardOutput = true,  // Capture output
                RedirectStandardError = true,   // Capture errors
                UseShellExecute = false,        // Required for redirection
                CreateNoWindow = true           // Don't show a window
            };

            List<string> outputLines = new List<string>();
            int exitCode = 0;

            using (var process = new Process())
            {
                process.StartInfo = processInfo;

                // Event handler for output data
                process.OutputDataReceived += (sender, e) =>
                {
                    if (!string.IsNullOrEmpty(e.Data))
                    {
                        outputLines.Add(e.Data);
                    }
                };

                // Event handler for error data
                process.ErrorDataReceived += (sender, e) =>
                {
                    if (!string.IsNullOrEmpty(e.Data))
                    {
                        outputLines.Add("ERROR: " + e.Data);
                    }
                };

                process.Start();

                // Begin asynchronous reading
                process.BeginOutputReadLine();
                process.BeginErrorReadLine();

                process.WaitForExit(); // Wait for process to complete
                exitCode = process.ExitCode;
            }

            return new ProcessResult { Output = outputLines.ToArray(), ExitCode = exitCode };
        }

        // Add this to your Main class
        private int GetWindowsBuildNumber()
        {
            // Basic approach using Environment.OSVersion
            var osVersion = Environment.OSVersion;
            var buildNumber = osVersion.Version.Build;

            return buildNumber;
        }


        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();
    }

    public class ProcessResult
    {
        public string[] Output { get; set; }
        public int ExitCode { get; set; }
    }
}
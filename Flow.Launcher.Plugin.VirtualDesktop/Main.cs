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
            int currentDesktopIndex = GetCurrentDesktopIndex();
            string currentDesktopName = desktops[currentDesktopIndex];

            // Handle rename command
            if (query.SearchTerms.Length > 0 && query.SearchTerms[0].Equals("rename", StringComparison.OrdinalIgnoreCase))
            {
                return HandleRenameDesktop(query, currentDesktopName);
            }

            // Handle new desktop command
            if (query.SearchTerms.Length > 0 && query.SearchTerms[0].Equals("new", StringComparison.OrdinalIgnoreCase))
            {
                return HandleCreateNewDesktop(query);
            }

            // Display current desktop info
            results.Add(new Result
            {
                Title = $"Current: {currentDesktopName}",
                SubTitle = "Type 'rename [name]' to rename this desktop or 'new [name]' to create a new desktop",
                IcoPath = IconPath,
                Action = (e) => false
            });

            // Add new desktop option in the main results
            results.Add(CreateNewDesktopResult());

            // List other desktops for switching
            for (int i = 0; i < desktops.Length; i++)
            {
                if (i == currentDesktopIndex)
                {
                    continue; // Skip the current desktop
                }

                results.Add(CreateSwitchDesktopResult(desktops[i], i));
            }

            return results;
        }

        private List<Result> HandleRenameDesktop(Query query, string currentDesktopName)
        {
            var results = new List<Result>();
            string newName = string.Join(" ", query.SearchTerms.Skip(1));

            if (!string.IsNullOrWhiteSpace(newName))
            {
                results.Add(new Result
                {
                    Title = $"Rename current desktop to: {newName}",
                    SubTitle = $"Current name: {currentDesktopName}",
                    IcoPath = IconPath,
                    Action = (e) =>
                    {
                        RenameCurrentDesktop(newName);
                        return true;
                    }
                });
            }
            else
            {
                results.Add(new Result
                {
                    Title = "Rename current desktop",
                    SubTitle = "Please enter a new name after the 'rename' command",
                    IcoPath = IconPath,
                    Action = (e) => false
                });
            }

            return results;
        }

        private List<Result> HandleCreateNewDesktop(Query query)
        {
            var results = new List<Result>();
            string newName = string.Join(" ", query.SearchTerms.Skip(1));

            if (!string.IsNullOrWhiteSpace(newName))
            {
                results.Add(new Result
                {
                    Title = $"Create new desktop with name: {newName}",
                    SubTitle = "Creates and switches to a new desktop with the specified name",
                    IcoPath = IconPath,
                    Action = (e) =>
                    {
                        CreateNewDesktopWithName(newName);
                        return true;
                    }
                });
            }
            else
            {
                results.Add(CreateNewDesktopResult());
            }

            return results;
        }

        private Result CreateNewDesktopResult()
        {
            return new Result
            {
                Title = "Create new desktop",
                SubTitle = "Creates and switches to a new desktop",
                IcoPath = IconPath,
                Action = (e) =>
                {
                    CreateNewDesktop();
                    return true;
                }
            };
        }

        private Result CreateSwitchDesktopResult(string desktopName, int desktopIndex)
        {
            return new Result
            {
                Title = desktopName,
                SubTitle = "Switch to this desktop",
                IcoPath = IconPath,
                Action = (e) =>
                {
                    SwitchToDesktop(desktopIndex);
                    return true;
                }
            };
        }

        private void RenameCurrentDesktop(string newName)
        {
            CallVDManager($"/GetCurrentDesktop /Name:{newName}");
        }

        private void CreateNewDesktop()
        {
            // Create a new desktop and get its index
            int newDesktopIndex = CallVDManager("/New").ExitCode;

            // Switch to the newly created desktop
            SwitchToDesktop(newDesktopIndex);
        }

        private void CreateNewDesktopWithName(string name)
        {
            // Create new desktop with name and get its index
            int newDesktopIndex = CallVDManager($"/New /Name:{name}").ExitCode;

            // If the desktop was created successfully, switch to it
            if (newDesktopIndex >= 0)
            {
                SwitchToDesktop(newDesktopIndex);
            }
        }

        private void SwitchToDesktop(int index)
        {
            CallVDManager($"/GetDesktop:{index} /MoveActiveWindow");
            CallVDManager($"/Switch:{index}");
        }

        private int GetCurrentDesktopIndex()
        {
            return CallVDManager("/Q /GetCurrentDesktop").ExitCode;
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
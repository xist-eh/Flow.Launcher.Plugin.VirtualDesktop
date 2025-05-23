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
        private const string IconPath = "Images/vd.png";
        private string vdExeName = "VirtualDesktop11.exe";
        private string vdFullPath;

        // Cache of desktop information to reduce process calls
        private string[] _cachedDesktops = Array.Empty<string>();
        private int _cachedCurrentDesktopIndex = -1;
        private DateTime _lastCacheUpdate = DateTime.MinValue;
        private readonly TimeSpan _cacheDuration = TimeSpan.FromSeconds(2); // Cache expires after 2 seconds
        private bool _stateChanged = false; // Flag to track if state has changed since last cache
        private int _lastSwitchedToDesktopIndex = -1; // Track the last switched desktop

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

            // Initially populate the cache
            RefreshDesktopCache();
        }

        public List<Result> Query(Query query)
        {
            var results = new List<Result>();

            if (query.SearchTerms.Any(s => s.Contains("debug", StringComparison.OrdinalIgnoreCase)))
            {
                results.Add(new Result { Title = "Build Number", SubTitle = GetWindowsBuildNumber().ToString(), IcoPath = IconPath });
                results.Add(new Result { Title = "VD Path", SubTitle = vdFullPath, IcoPath = IconPath });
            }

            // Only refresh the cache if it's expired or state has changed
            if (IsCacheExpired() || _stateChanged)
            {
                RefreshDesktopCache();
                _stateChanged = false;
            }

            string[] desktops = _cachedDesktops;
            int currentDesktopIndex = _cachedCurrentDesktopIndex;

            // Ensure we have valid data
            if (desktops.Length == 0 || currentDesktopIndex < 0 || currentDesktopIndex >= desktops.Length)
            {
                // Something is wrong with our cache, force a refresh
                RefreshDesktopCache();
                desktops = _cachedDesktops;
                currentDesktopIndex = _cachedCurrentDesktopIndex;

                // If still invalid, return an error message
                if (desktops.Length == 0 || currentDesktopIndex < 0 || currentDesktopIndex >= desktops.Length)
                {
                    results.Add(new Result
                    {
                        Title = "Error retrieving virtual desktops",
                        SubTitle = "Please check if the virtual desktop manager is working correctly",
                        IcoPath = IconPath,
                        Action = (e) => false
                    });
                    return results;
                }
            }

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

            // Handle move window command
            if (query.SearchTerms.Length > 0 && query.SearchTerms[0].Equals("move", StringComparison.OrdinalIgnoreCase))
            {
                // Show options for moving the active window to other desktops
                return HandleMoveWindowToDesktop(query, desktops, currentDesktopIndex);
            }

            // If search term exists but isn't a command, treat as fuzzy search
            if (query.SearchTerms.Length > 0 &&
                !query.SearchTerms[0].Equals("new", StringComparison.OrdinalIgnoreCase) &&
                !query.SearchTerms[0].Equals("rename", StringComparison.OrdinalIgnoreCase) &&
                !query.SearchTerms[0].Equals("move", StringComparison.OrdinalIgnoreCase))
            {
                return HandleFuzzyDesktopSearch(query, desktops, currentDesktopIndex);
            }

            // List available desktops for switching (excluding current one)

            // First, add the last switched desktop at the top if it's valid and not the current one
            if (_lastSwitchedToDesktopIndex >= 0 &&
                _lastSwitchedToDesktopIndex < desktops.Length &&
                _lastSwitchedToDesktopIndex != currentDesktopIndex)
            {
                var lastDesktopResult = CreateSwitchDesktopResult(desktops[_lastSwitchedToDesktopIndex], _lastSwitchedToDesktopIndex);
                lastDesktopResult.Title = $"↩ {desktops[_lastSwitchedToDesktopIndex]}"; // Add a return arrow to indicate "last used"
                lastDesktopResult.SubTitle = "Switch back to previously used desktop";
                results.Add(lastDesktopResult);
            }

            // Then add all other desktops
            for (int i = 0; i < desktops.Length; i++)
            {
                if (i == currentDesktopIndex || i == _lastSwitchedToDesktopIndex)
                {
                    continue; // Skip the current desktop and last switched desktop (already added)
                }

                results.Add(CreateSwitchDesktopResult(desktops[i], i));
            }

            // Add new desktop option
            results.Add(CreateNewDesktopResult());

            // Display current desktop info with additional move command hint
            results.Add(new Result
            {
                Title = $"Current: {currentDesktopName}",
                SubTitle = "Type 'rename [name]', 'new [name]', or 'move' to move window",
                IcoPath = IconPath,
                Action = (e) => false
            });

            return results;
        }

        private bool IsCacheExpired()
        {
            return (DateTime.Now - _lastCacheUpdate) > _cacheDuration;
        }

        private void RefreshDesktopCache()
        {
            string[] desktops = GetAllDesktopsFromSystem();
            int currentIndex = GetCurrentDesktopIndexFromSystem();

            // Only update the cache if we got valid data
            if (desktops.Length > 0 && currentIndex >= 0 && currentIndex < desktops.Length)
            {
                _cachedDesktops = desktops;
                _cachedCurrentDesktopIndex = currentIndex;
                _lastCacheUpdate = DateTime.Now;
            }
        }

        private void MarkStateChanged()
        {
            _stateChanged = true;
        }

        private List<Result> HandleFuzzyDesktopSearch(Query query, string[] desktops, int currentDesktopIndex)
        {
            var results = new List<Result>();
            string searchTerm = string.Join(" ", query.SearchTerms).ToLower();

            // Custom fuzzy search implementation
            var matches = new List<(int index, string desktop, int score)>();
            for (int i = 0; i < desktops.Length; i++)
            {
                if (i != currentDesktopIndex) // Skip current desktop
                {
                    string desktopName = desktops[i].ToLower();
                    int score = CalculateFuzzyMatchScore(searchTerm, desktopName);
                    if (score >= 70) // Only include matches with score >= 70
                    {
                        matches.Add((i, desktops[i], score));
                    }
                }
            }

            // Sort by match score in descending order
            matches.Sort((a, b) => b.score.CompareTo(a.score));

            // Add desktop results
            foreach (var match in matches)
            {
                results.Add(new Result
                {
                    Title = match.desktop,
                    SubTitle = "Switch to this desktop",
                    IcoPath = IconPath,
                    Score = match.score,
                    Action = (e) =>
                    {
                        SwitchToDesktop(match.index);
                        return true;
                    }
                });
            }

            // If no results, provide feedback
            if (matches.Count == 0)
            {
                results.Add(new Result
                {
                    Title = "No matching desktops found",
                    SubTitle = "Try a different search term or create a new desktop",
                    IcoPath = IconPath,
                    Action = (e) => false
                });

                // Add option to create a new desktop with this name
                results.Add(new Result
                {
                    Title = $"Create new desktop named: '{searchTerm}'",
                    SubTitle = "Creates and switches to a new desktop with this name",
                    IcoPath = IconPath,
                    Action = (e) =>
                    {
                        CreateNewDesktopWithName(searchTerm);
                        return true;
                    }
                });
            }

            return results;
        }

        // Custom fuzzy search scoring function
        private int CalculateFuzzyMatchScore(string query, string target)
        {
            if (string.IsNullOrEmpty(query))
                return 0;

            return _context.API.FuzzySearch(query, target).Score;
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
            MarkStateChanged();
        }

        private void CreateNewDesktop()
        {
            // Create a new desktop and get its index
            int newDesktopIndex = CallVDManager("/New").ExitCode;

            // If the desktop was created successfully, switch to it
            if (newDesktopIndex >= 0)
            {
                SwitchToDesktop(newDesktopIndex);
                MarkStateChanged();
            }
        }

        private void CreateNewDesktopWithName(string name)
        {
            // Create new desktop with name and get its index
            int newDesktopIndex = CallVDManager($"/New /Name:{name}").ExitCode;

            // If the desktop was created successfully, switch to it
            if (newDesktopIndex >= 0)
            {
                SwitchToDesktop(newDesktopIndex);
                MarkStateChanged();
            }
        }

        private void SwitchToDesktop(int index)
        {
            // Store current desktop as the last switched-from desktop
            if (_cachedCurrentDesktopIndex >= 0 && _cachedCurrentDesktopIndex != index)
            {
                _lastSwitchedToDesktopIndex = _cachedCurrentDesktopIndex;
            }

            // Combined call to both move active window and switch desktop
            // to make a single external process call instead of two
            CallVDManager($"/GetDesktop:{index} /MoveActiveWindow /Switch");

            // Force an immediate cache refresh to reflect the new desktop state
            // This ensures the next time Flow is opened the current desktop is correct
            MarkStateChanged();

            // Small delay to ensure the desktop switch completes before refreshing
            Thread.Sleep(100);

            // Immediately refresh the cache so any subsequent operations show the correct desktop
            RefreshDesktopCache();
        }

        // These methods interact directly with the system, bypassing the cache
        private int GetCurrentDesktopIndexFromSystem()
        {
            return CallVDManager("/Q /GetCurrentDesktop").ExitCode;
        }

        private string[] GetAllDesktopsFromSystem()
        {
            return CallVDManager("/Q /List").Output;
        }

        private void MoveActiveWindowToDesktop(int targetDesktopIndex)
        {
            // Get the current desktop's window list
            var result = CallVDManager("/q /GetCurrentDesktop /ListWindowsOnDesktop");

            // Check if we have enough windows in the output
            if (result.Output != null && result.Output.Length > 1)
            {
                // Get the second window from the list
                string windowHandle = result.Output[1];

                // Move the window to the target desktop
                CallVDManager($"/GetDesktop:{targetDesktopIndex} /MoveWindowHandle:{windowHandle}");

                // Mark that state has changed
                MarkStateChanged();
            }
        }

        // Add a UI method to create results for moving windows to other desktops
        private Result CreateMoveWindowToDesktopResult(string desktopName, int desktopIndex)
        {
            return new Result
            {
                Title = $"Move window to {desktopName}",
                SubTitle = "Moves the active window to this desktop without switching",
                IcoPath = IconPath,
                Action = (e) =>
                {
                    MoveActiveWindowToDesktop(desktopIndex);
                    return true;
                }
            };
        }

        // Add this new method to handle the move window command
        private List<Result> HandleMoveWindowToDesktop(Query query, string[] desktops, int currentDesktopIndex)
        {
            var results = new List<Result>();

            // If there are additional search terms, use them for fuzzy searching desktops
            if (query.SearchTerms.Length > 1)
            {
                string searchTerm = string.Join(" ", query.SearchTerms.Skip(1)).ToLower();

                // Filter desktops by name (fuzzy match)
                var matches = new List<(int index, string desktop, int score)>();
                for (int i = 0; i < desktops.Length; i++)
                {
                    if (i != currentDesktopIndex) // Skip current desktop
                    {
                        string desktopName = desktops[i].ToLower();
                        int score = CalculateFuzzyMatchScore(searchTerm, desktopName);
                        if (score >= 70) // Only include matches with score >= 70
                        {
                            matches.Add((i, desktops[i], score));
                        }
                    }
                }

                // Sort by match score in descending order
                matches.Sort((a, b) => b.score.CompareTo(a.score));

                // Add desktop results for moving window
                foreach (var match in matches)
                {
                    results.Add(CreateMoveWindowToDesktopResult(match.desktop, match.index));
                }

                // If no results, provide feedback
                if (matches.Count == 0)
                {
                    results.Add(new Result
                    {
                        Title = "No matching desktops found",
                        SubTitle = "Try a different search term",
                        IcoPath = IconPath,
                        Action = (e) => false
                    });
                }
            }
            else
            {
                // List all available desktops for moving the window (excluding current one)
                for (int i = 0; i < desktops.Length; i++)
                {
                    if (i != currentDesktopIndex) // Skip current desktop
                    {
                        results.Add(CreateMoveWindowToDesktopResult(desktops[i], i));
                    }
                }

                // Provide instructions
                results.Add(new Result
                {
                    Title = "Move active window",
                    SubTitle = "Select a desktop to move the active window to, without switching",
                    IcoPath = IconPath,
                    Action = (e) => false
                });
            }

            return results;
        }

        private ProcessResult CallVDManager(string args)
        {
            // Setup process configuration
            var processInfo = new ProcessStartInfo
            {
                FileName = vdFullPath,
                Arguments = args,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            List<string> outputLines = new List<string>();
            int exitCode = 0;

            // Execute process and capture output
            using (var process = new Process())
            {
                process.StartInfo = processInfo;

                // Setup output capture
                process.OutputDataReceived += (sender, e) =>
                {
                    if (!string.IsNullOrEmpty(e.Data))
                    {
                        outputLines.Add(e.Data);
                    }
                };

                // Setup error capture
                process.ErrorDataReceived += (sender, e) =>
                {
                    if (!string.IsNullOrEmpty(e.Data))
                    {
                        outputLines.Add("ERROR: " + e.Data);
                    }
                };

                // Start the process
                process.Start();

                // Begin asynchronous reading
                process.BeginOutputReadLine();
                process.BeginErrorReadLine();

                // Wait for completion
                process.WaitForExit();
                exitCode = process.ExitCode;
            }

            return new ProcessResult
            {
                Output = outputLines.ToArray(),
                ExitCode = exitCode
            };
        }

        // Add this to your Main class
        static private int GetWindowsBuildNumber()
        {
            // Basic approach using Environment.OSVersion
            var osVersion = Environment.OSVersion;
            var buildNumber = osVersion.Version.Build;

            return buildNumber;
        }
    }

    public class ProcessResult
    {
        public string[] Output { get; set; }
        public int ExitCode { get; set; }
    }
}
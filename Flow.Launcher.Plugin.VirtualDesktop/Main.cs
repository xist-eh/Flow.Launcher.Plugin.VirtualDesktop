using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using Flow.Launcher.Plugin;
using WindowsDesktop;

namespace Flow.Launcher.Plugin.VirtualDesktop
{
    public class Main : IPlugin
    {
        private PluginInitContext _context;
        private const string IconPath = "Images/icon.png";

        public void Init(PluginInitContext context)
        {
            _context = context;
        }

        public List<Result> Query(Query query)
        {
            var results = new List<Result>();

            try
            {
                // Get all virtual desktops
                var desktops = WindowsDesktop.VirtualDesktop.GetDesktops();
                var currentDesktop = WindowsDesktop.VirtualDesktop.Current;

                // Add result for each desktop
                for (int i = 0; i < desktops.Length; i++)
                {
                    var desktop = desktops[i];
                    var desktopName = string.IsNullOrEmpty(desktop.Name) ? $"Desktop {i + 1}" : desktop.Name;
                    var isCurrent = desktop.Id == currentDesktop.Id;
                    var desktopIndex = i; // Capture the index for use in lambda

                    results.Add(new Result
                    {
                        Title = desktopName + (isCurrent ? " (Current)" : ""),
                        SubTitle = $"Switch to {desktopName}",
                        IcoPath = IconPath,
                        Score = isCurrent ? 100 - i : 90 - i,
                        Action = _ =>
                        {
                            return ExecuteStaThread(() =>
                            {
                                if (!isCurrent)
                                {
                                    var desktopToSwitch = WindowsDesktop.VirtualDesktop.GetDesktops()[desktopIndex];
                                    desktopToSwitch.Switch();
                                }
                                return true;
                            });
                        }
                    });
                }

                // Add option to create new desktop
                results.Add(new Result
                {
                    Title = "Create new virtual desktop",
                    SubTitle = "Create and switch to a new virtual desktop",
                    IcoPath = IconPath,
                    Score = 50,
                    Action = _ =>
                    {
                        return ExecuteStaThread(() =>
                        {
                            var newDesktop = WindowsDesktop.VirtualDesktop.Create();
                            newDesktop.Switch();
                            return true;
                        });
                    }
                });

                // Add option to remove current desktop
                if (desktops.Length > 1)
                {
                    results.Add(new Result
                    {
                        Title = "Remove current virtual desktop",
                        SubTitle = $"Remove the current desktop ({(string.IsNullOrEmpty(currentDesktop.Name) ? "Desktop" : currentDesktop.Name)})",
                        IcoPath = IconPath,
                        Score = 40,
                        Action = _ =>
                        {
                            return ExecuteStaThread(() =>
                            {
                                var current = WindowsDesktop.VirtualDesktop.Current;
                                current.Remove();
                                return true;
                            });
                        }
                    });
                }

                // Add option to move active window to another desktop
                if (desktops.Length > 1)
                {
                    foreach (var targetDesktop in desktops)
                    {
                        if (targetDesktop.Id == currentDesktop.Id)
                            continue;

                        var targetName = string.IsNullOrEmpty(targetDesktop.Name)
                            ? $"Desktop {Array.IndexOf(desktops, targetDesktop) + 1}"
                            : targetDesktop.Name;

                        var targetIndex = Array.IndexOf(desktops, targetDesktop); // Capture for use in lambda

                        results.Add(new Result
                        {
                            Title = $"Move window to {targetName}",
                            SubTitle = $"Move active window to {targetName}",
                            IcoPath = IconPath,
                            Score = 25,
                            Action = _ =>
                            {
                                return ExecuteStaThread(() =>
                                {
                                    try
                                    {

                                    }
                                    catch (Exception ex)
                                    {

                                    }
                                    return true;
                                });
                            }
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                results.Add(new Result
                {
                    Title = "Error accessing Virtual Desktops",
                    SubTitle = ex.Message,
                    IcoPath = IconPath,
                    Score = 100
                });
            }

            return results;
        }

        // Helper method to execute code in an STA thread
        private bool ExecuteStaThread(Func<bool> action)
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
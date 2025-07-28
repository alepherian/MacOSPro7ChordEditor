using System;
using System.IO;
using System.Runtime.InteropServices;

namespace Pro7ChordEditor
{
    public static class PlatformHelper
    {
        /// <summary>
        /// Gets the ProPresenter base folder path for the current platform
        /// Note: Libraries are in a "Libraries" subfolder within this path
        /// Windows: C:\Users\%UserProfile%\Documents\ProPresenter\Libraries
        /// macOS: ~/Documents/ProPresenter/Libraries
        /// </summary>
        public static string GetProPresenterPath()
        {
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            {
                // Windows path logic (your existing code)
                string customPathSettingsFilePath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "RenewedVision", "ProPresenter", "PathSettings.proPaths");
                
                if (File.Exists(customPathSettingsFilePath))
                {
                    try
                    {
                        using (var sr = new StreamReader(customPathSettingsFilePath))
                        {
                            return System.Text.RegularExpressions.Regex.Match(
                                sr.ReadToEnd().Replace(@"\\", @"\"), 
                                @"(?<=Base=).*(?=;)").Value;
                        }
                    }
                    catch (IOException)
                    {
                        // Fall back to default
                    }
                }
                
                return Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.Personal), 
                    "ProPresenter");
            }
            else if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
            {
                // macOS path logic
                string homeDir = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
                
                // Check for custom path settings first
                string customPathSettingsFilePath = Path.Combine(
                    homeDir, "Library", "Application Support", "RenewedVision", "ProPresenter", "PathSettings.proPaths");
                
                if (File.Exists(customPathSettingsFilePath))
                {
                    try
                    {
                        using (var sr = new StreamReader(customPathSettingsFilePath))
                        {
                            var content = sr.ReadToEnd();
                            var match = System.Text.RegularExpressions.Regex.Match(content, @"(?<=Base=).*(?=;)");
                            if (match.Success && !string.IsNullOrEmpty(match.Value))
                            {
                                return match.Value;
                            }
                        }
                    }
                    catch (IOException)
                    {
                        // Fall back to checking common paths
                    }
                }
                
                // Check common macOS locations for ProPresenter
                string[] possiblePaths = {
                    Path.Combine(homeDir, "Documents", "ProPresenter"),
                    Path.Combine(homeDir, "Library", "Application Support", "RenewedVision", "ProPresenter"),
                    "/Applications/ProPresenter.app/Contents/Resources/ProPresenter"
                };
                
                foreach (string path in possiblePaths)
                {
                    if (Directory.Exists(path))
                    {
                        return path;
                    }
                }
                
                // Default to Documents/ProPresenter (same as Windows structure)
                return Path.Combine(homeDir, "Documents", "ProPresenter");
            }
            else
            {
                // Linux or other Unix-like systems
                string homeDir = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
                return Path.Combine(homeDir, "ProPresenter");
            }
        }
        
        public static string GetLogPath()
        {
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            {
                return Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) ?? 
                       Path.GetTempPath();
            }
            else if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
            {
                // macOS: Use Application Support directory
                string homeDir = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
                string logDir = Path.Combine(homeDir, "Library", "Application Support", "Pro7ChordEditor");
                Directory.CreateDirectory(logDir); // Ensure directory exists
                return logDir;
            }
            else
            {
                // Linux: Use ~/.local/share
                string homeDir = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
                string logDir = Path.Combine(homeDir, ".local", "share", "Pro7ChordEditor");
                Directory.CreateDirectory(logDir);
                return logDir;
            }
        }
        
        public static string NormalizePath(string path)
        {
            // Convert Windows-style paths to cross-platform
            return path.Replace('\\', Path.DirectorySeparatorChar);
        }
        
        public static string GetProPresenterProcessName()
        {
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            {
                return "ProPresenter";
            }
            else if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
            {
                return "ProPresenter"; // May need to check actual process name on macOS
            }
            else
            {
                return "ProPresenter";
            }
        }
    }
}
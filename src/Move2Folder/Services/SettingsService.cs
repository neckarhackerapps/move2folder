using System;
using System.IO;
using System.Text.Json;

namespace Move2Folder.Services
{
    public class UserSettings
    {
        public string LastFolderEntryId { get; set; } = string.Empty;
        public string LastFolderStoreId { get; set; } = string.Empty;
        public string LastFolderPath { get; set; } = string.Empty;
        public double WindowWidth { get; set; } = 550;
        public double WindowHeight { get; set; } = 600;
    }

    public static class SettingsService
    {
        private static readonly string SettingsDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "Move2Folder");

        private static readonly string SettingsFile = Path.Combine(SettingsDir, "settings.json");

        public static UserSettings Load()
        {
            try
            {
                if (File.Exists(SettingsFile))
                {
                    string json = File.ReadAllText(SettingsFile);
                    var settings = JsonSerializer.Deserialize<UserSettings>(json);
                    if (settings != null)
                        return settings;
                }
            }
            catch
            {
                // Fall back to default
            }

            return new UserSettings();
        }

        public static void Save(UserSettings settings)
        {
            try
            {
                if (!Directory.Exists(SettingsDir))
                {
                    Directory.CreateDirectory(SettingsDir);
                }

                string json = JsonSerializer.Serialize(settings, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(SettingsFile, json);
            }
            catch
            {
                // Suppress save errors silently to not interrupt the user
            }
        }
    }
}

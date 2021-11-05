using Microsoft.Extensions.Configuration;
using System;

namespace HIGKnowledgePortal
{
    public class Configuration
    {
        public static IConfiguration ConfigFile => GetConfigurationFile("appSettings.json");
        public static string MappingFilePath => (string)GetConfigValue("filePath");
        public static string GraphUrl => (string)GetConfigValue("graphUrl");
        public static string SharePointOrg => (string)GetConfigValue("spTenant");
        public static string SiteName => (string)GetConfigValue("spSite");
        public static string ListName => (string)GetConfigValue("spList");
        public static string DriveName => (string)GetConfigValue("driveName");
        public static string DownloadDirectory => (string)GetConfigValue("downloadDirectory");
        public static string AuthTenant => (string)GetConfigValue("tenant", "authentication");
        public static string AuthClient => (string)GetConfigValue("client_id", "authentication");
        public static string AuthClientSecret => (string)GetConfigValue("client_secret", "authentication");
        public static string AuthToken { get; set; }

        /// <summary>
        /// Get the configuration file appSettings.json
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        private static IConfiguration GetConfigurationFile(string fileName)
        {
            IConfiguration config = new ConfigurationBuilder()
                .AddJsonFile(fileName)
                .Build();

            return config;
        }

        /// <summary>
        /// Gets a configuration value from the config file
        /// </summary>
        /// <param name="name"></param>
        /// <param name="typeCode"></param>
        /// <returns></returns>
        private static object GetConfigValue(string name, TypeCode typeCode = TypeCode.String)
        {
            Type configType = GetType(typeCode);
            var configValue = ConfigFile.GetValue(configType, name);
            return configValue;
        }

        /// <summary>
        /// Gets a configuration value from a section in the config file
        /// </summary>
        /// <param name="name"></param>
        /// <param name="section"></param>
        /// <param name="typeCode"></param>
        /// <returns></returns>
        private static object GetConfigValue(string name, string section, TypeCode typeCode = TypeCode.String)
        {
            Type configType = GetType(typeCode);
            var configSection = ConfigFile.GetSection(section);
            var configValue = configSection.GetValue(configType, name);
            return configValue;
        }

        /// <summary>
        /// Get a Type based on the TypeCode
        /// </summary>
        /// <param name="typeCode"></param>
        /// <returns></returns>
        private static Type GetType(TypeCode typeCode)
        {
            Type type;

            switch (typeCode)
            {
                case TypeCode.String:
                    type = typeof(string);
                    break;
                case TypeCode.Boolean:
                    type = typeof(bool);
                    break;
                case TypeCode.Int32:
                    type = typeof(int);
                    break;
                default:
                    type = typeof(string);
                    break;
            }

            return type;
        }
    }
}

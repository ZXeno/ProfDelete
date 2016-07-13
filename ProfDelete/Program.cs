using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Management;
using System.Net.NetworkInformation;
using System.Threading.Tasks;

namespace ProfDelete
{
    class Program
    {
        private static int _olderThan = 14;
        private static bool _doPerformanceTiming = false;
        private static bool _isSilent = false;
        private static string _device = "";
        private static string _rootNamespace = "root\\cimv2";
        private static string _wmiQuery = "SELECT * FROM Win32_UserProfile WHERE LocalPath LIKE '%Users%'";

        static void Main(string[] args)
        {
            WriteLine("Old Profile Delete Tool");
            WriteLine("Copyright: GPLv3");
            WriteLine("Written by: Jonathan Cain");
            WriteLine();

            if (args.Length > 0 && !ValidateAndParseArgs(args.ToList()))
            {
                Console.WriteLine("Press ENTER to exit.");
                Console.ReadLine();
                return;
            }

            if (args.Length == 0)
            {
                if (!RequestArgs())
                {
                    Console.WriteLine("Press ENTER to exit.");
                    Console.ReadLine();
                    return;
                }
            }

            // preventing actual execution during testing
            WriteLine($"Runing profile cleanup on system: {_device}");
            WriteLine($"Profiles older than: {_olderThan}");
            WriteLine($"Performance meansurement: {_doPerformanceTiming}");


            var scope = new ManagementScope($"\\\\{_device}\\{_rootNamespace}");
            scope.Connect();
            var objQuery = new ObjectQuery(_wmiQuery);
            
            int totalProfiles = 0;
            long slowest = long.MinValue;
            string slowestPath = "";
            long fastest = long.MaxValue;
            string fastestPath = "";


            var totalStopwatch = new Stopwatch();
            if (_doPerformanceTiming)
            {
                totalStopwatch.Start();
            }
            
            try
            {
                var searcher = new ManagementObjectSearcher(scope, objQuery);
                var colItems = searcher.Get();


                var profiles = new List<ManagementObject>();


                foreach (ManagementObject queryObj in colItems)
                {
                    var queryObjDate = queryObj?["LastUseTime"];
                    if (queryObjDate == null) { continue; }


                    var date = ManagementDateTimeConverter.ToDateTime(queryObjDate.ToString());


                    if (DateTime.Now.DayOfYear - date.DayOfYear >= _olderThan)
                    {
                        profiles.Add(queryObj);
                    }
                }


                var countText = $"{profiles.Count} profile{(profiles.Count == 1 ? String.Empty : "s")} slated for deletion.";
                totalProfiles = profiles.Count;
                
                WriteLine($"Deleting profiles older than {_olderThan}");
                WriteLine(countText);

                try
                {
                    Parallel.ForEach(profiles, (queryObj) =>
                    {
                        var stopWatch = new Stopwatch();

                        var userProfName = queryObj["LocalPath"].ToString().Split('\\').Last();
                        var path = $"\\\\{_device}\\C$\\Users\\{userProfName}";
                        WriteLine($"Begin delete of {path}.");

                        if (_doPerformanceTiming)
                        {
                            stopWatch.Start();
                        }

                        try
                        {
                            queryObj.Delete();
                            queryObj.Dispose();
                        }
                        catch (Exception e)
                        {
                            var prevColor = Console.ForegroundColor;
                            Console.ForegroundColor = ConsoleColor.Red;
                            WriteLine($"There was an error deleting profile {userProfName}.");
                            WriteLine($"[ERROR]: {e.Message}");

                            if (e.InnerException != null)
                            {
                                WriteLine($"[INNER EXCEPTION MESSAGE] {e.InnerException.Message}");
                            }
                            Console.ForegroundColor = prevColor;
                        }
                        
                        

                        if (_doPerformanceTiming)
                        {
                            stopWatch.Stop();


                            if (stopWatch.ElapsedMilliseconds > slowest)
                            {
                                slowest = stopWatch.ElapsedMilliseconds;
                                slowestPath = path.ToString();
                            }
                            if (stopWatch.ElapsedMilliseconds < fastest)
                            {
                                fastest = stopWatch.ElapsedMilliseconds;
                                fastestPath = path.ToString();
                            }
                        }

                        WriteLine($"Delete {path} done.");

                        if (_doPerformanceTiming)
                        {
                            WriteLine($"Took {stopWatch.ElapsedMilliseconds} ms || {stopWatch.ElapsedMilliseconds / 1000} seconds.");
                        }
                    });
                }
                catch (Exception e)
                {
                    var prevColor = Console.ForegroundColor;
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.Error.WriteLine(e.Message);
                    Console.ForegroundColor = prevColor;
                }

            }
            catch (Exception e)
            {
                var prevColor = Console.ForegroundColor;
                Console.ForegroundColor = ConsoleColor.Red;
                Console.Error.WriteLine($"An error occurred while querying WMI: {e.Message}");
                WriteLine("Query - " + _wmiQuery);
                Console.ForegroundColor = prevColor;
            }

            if (_doPerformanceTiming)
            {
                totalStopwatch.Stop();


                WriteLine();
                WriteLine($"Fastest     : {fastestPath}");
                WriteLine($"Time in ms  : {fastest} ms");
                WriteLine($"Time in s   : {fastest/1000} seconds.");
                WriteLine();
                WriteLine($"Slowest     : {slowestPath}");
                WriteLine($"Time in ms  : {slowest} ms");
                WriteLine($"Time in s   : {slowest/1000} seconds.");
                WriteLine();
                WriteLine($"Total profiles deleted: {totalProfiles}");
                WriteLine($"Total Time  : {totalStopwatch.ElapsedMilliseconds} ms");
                WriteLine($"Total Time  : {totalStopwatch.ElapsedMilliseconds/1000} seconds.");
                WriteLine();
            }

            WriteLine("Done.");
            Console.ReadLine();
        }

        private static bool ValidateAndParseArgs(List<string> args)
        {
            if (args.Count == 0) { return false; }

            var deviceArgMatch = args.FirstOrDefault(arg => arg.Contains("/device:"));
            if (string.IsNullOrWhiteSpace(deviceArgMatch))
            {
                WriteLine("No device specified!");
                return false;
            }

            _device = deviceArgMatch.Split(':').Last();

            if (!TestConnectivity(_device))
            {
                var prevColor = Console.ForegroundColor;
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"Unable to connect to device {_device}.");
                Console.ForegroundColor = prevColor;
                return false;
            }

            var daysArgMatch = args.FirstOrDefault(arg => arg.Contains("/days:"));
            if (!string.IsNullOrWhiteSpace(daysArgMatch))
            {
                int count = -1;
                try
                {
                    var days = daysArgMatch.Split(':').Last();
                    count = int.Parse(days);
                }
                catch (Exception) { }
                

                if (count <= 0)
                {
                    var prevColor = Console.ForegroundColor;
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("Unable to parse the \"Days Older Than\" parameter. Cancelling...");
                    Console.ForegroundColor = prevColor;
                    return false;
                }

                if (count >= 365)
                {
                    var prevColor = Console.ForegroundColor;
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("\"Days Older Than\" parameter was set to a value greater than 365. Cancelling.");
                    Console.ForegroundColor = prevColor;
                    return false;
                }

                _olderThan = count;
            }
            else
            {
                var prevColor = Console.ForegroundColor;
                Console.ForegroundColor = ConsoleColor.Yellow;
                WriteLine($"No \"Days Older Than\" parameter specified. Using default of {_olderThan} days.");
                Console.ForegroundColor = prevColor;
            }

            if (args.Contains("/stopwatch"))
            {
                _doPerformanceTiming = true;
            }

            if (args.Contains("/quiet") || args.Contains("/q"))
            {
                _isSilent = true;
            }

            return true;
        }

        private static void WriteLine(string line = "")
        {
            if (_isSilent) { return; }

            Console.WriteLine(line);
        }

        private static bool RequestArgs()
        {
            Console.WriteLine("Device ID:");
            _device = Console.ReadLine();

            if (!TestConnectivity(_device))
            {
                Console.WriteLine($"Could not connect to device {_device}");
                return false;
            }

            Console.WriteLine();
            Console.WriteLine($"Delete profiles older than (1-365): (Default is {_olderThan})");
            var daysEntered = Console.ReadLine();
            int parseAttempt = 0;

            if (int.TryParse(daysEntered, out parseAttempt))
            {
                if (parseAttempt <= 0)
                {
                    var prevColor = Console.ForegroundColor;
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine("Minimum time is 1 day. Set to delete profiles older than 1 day.");
                    _olderThan = 1;
                    Console.ForegroundColor = prevColor;
                }
                if (parseAttempt > 365)
                {
                    var prevColor = Console.ForegroundColor;
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine("Maximum time is 365 days. Set to delete profiles older than 365 days.");
                    _olderThan = 365;
                    Console.ForegroundColor = prevColor;
                }
            }
            else
            {
                Console.WriteLine($"Unable to parse the \"days older than\" argument.\nDo you want to use the default ({_olderThan})? (y/yes || n/no)");
                var result = Console.ReadLine()?.ToLower();
                if (!string.IsNullOrWhiteSpace(result) && !(result == "y" || result == "yes"))
                {
                    return false;
                }
            }

            Console.WriteLine("Do you want view performance stats?  (y/yes || n/no)");
            var perfResult = Console.ReadLine()?.ToLower();
            if (!string.IsNullOrWhiteSpace(perfResult) && (perfResult == "y" || perfResult == "yes"))
            {
                _doPerformanceTiming = true;
            }

            return true;
        }

        private static bool TestConnectivity(string device)
        {
            try
            {
                var reply = new Ping().Send(device, 3000);

                return reply?.Status == IPStatus.Success;
            }
            catch (Exception e)
            {
                return false;
            }
        }
    }
}
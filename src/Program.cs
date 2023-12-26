using CommandLine;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Timers;

namespace SignCBS
{
    class Program
    {
        internal class Options
        {
            [Option('r', "recurse", HelpText = "Process all folders in the given path", Required = false, Default = false)]
            public bool Recursive { get; set; }

            [Option('p', "path", HelpText = @"Path containing CBS Cab packages.", Required = true)]
            public string cbspath { get; set; }

            [Option('l', "logfile", HelpText = "Log output to new file", Required = false, Default = false)]
            public bool Logger { get; set; }


        }

        [StructLayout(LayoutKind.Sequential)]
        public struct SYSTEMTIME
        {
            public short wYear;
            public short wMonth;
            public short wDayOfWeek;
            public short wDay;
            public short wHour;
            public short wMinute;
            public short wSecond;
            public short wMilliseconds;
        }


        //Timer timer;
        static DateTime tempDateTime;
        static int ErrorCount;
        static int FileCount;
        static int CurrentCount;
        static string Logger;
        static DateTime savedDate;
        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern bool SetSystemTime(ref SYSTEMTIME st);
        static void Main(string[] args)
        {
            try
            {
                Console.Title = "SignCBS";
                Console.WriteLine("Windows Phone CBS Package Signer");


                Parser.Default.ParseArguments<Options>(args).WithParsed(m =>
                {
                    if (args.Length != 0)
                    {
                        Console.WriteLine(System.Reflection.Assembly.GetExecutingAssembly().GetName().Version + Environment.NewLine);
                    }
                    // not impl yet //
                    //bool debug = m.Verbose;
                    bool certsInstalled = false;
                    tempDateTime = DateTime.Now;


                    Console.WriteLine("[Date needs to be changed to allow OEM Certificate to be valid]");
                    savedDate = DateTime.Now;
                    SYSTEMTIME st = new SYSTEMTIME();
                    // All of these must be short
                    st.wYear = (short)2015;
                    st.wMonth = (short)1;
                    st.wDay = (short)1;
                    st.wHour = (short)DateTime.Now.Hour;
                    st.wMinute = (short)DateTime.Now.Minute;
                    st.wSecond = (short)DateTime.Now.Second;

                    // invoke the SetSystemTime method now
                    SetSystemTime(ref st);

                    TimeService(false);


                    Console.WriteLine("[Setting environment variables for SignTool]");
                    Environment.SetEnvironmentVariable("SIGN_OEM", "1");
                    Environment.SetEnvironmentVariable("SIGN_WITH_TIMESTAMP", "1");

                    Task.Delay(500);

                    if (File.Exists(@"certs.installed") == false)
                    {
                        Console.WriteLine("[Installing OEM Certs, please wait.]");
                        var oemCerts = new Process();

                        oemCerts.StartInfo.FileName = @".\Tools\installoemcerts.bat";
                        oemCerts.StartInfo.RedirectStandardOutput = true;
                        oemCerts.StartInfo.RedirectStandardError = true;
                        oemCerts.StartInfo.UseShellExecute = false;
                        oemCerts.StartInfo.CreateNoWindow = true;
                        oemCerts.OutputDataReceived += Signcmd_OutputDataReceived;
                        oemCerts.ErrorDataReceived += Signcmd_ErrorDataReceived;
                        oemCerts.Start();
                        oemCerts.BeginOutputReadLine();
                        oemCerts.BeginErrorReadLine();
                        oemCerts.WaitForExit();
                        //oemCerts.Kill();
                        File.WriteAllText("certs.installed", "OEM Certificates Installed");
                        certsInstalled = true;
                    }
                    else
                    {
                        certsInstalled = true;
                    }

                    while (certsInstalled == false)
                    {
                        Task.Delay(500);
                    }

                    CurrentCount = 0;
                    Console.WriteLine("[Signing Packages]");

                    IEnumerable<string> EnumFolders;

                    if (m.Recursive == true)
                    {
                        EnumFolders = Directory.EnumerateDirectories(m.cbspath, "*", SearchOption.AllDirectories);
                        RecursiveSign(EnumFolders);
                    }
                    else
                    {
                        SingularSign(m.cbspath);
                    }



                    TimeService(true);

                    SetRealTime();

                    if (ErrorCount != 0)
                    {
                        Console.WriteLine("Errors occured during Signing, please check log above");
                        if (m.Logger == true)
                        {
                            File.WriteAllText("OUTPUT.LOG", Logger);
                        }
                    }
                    else
                    {
                        Console.WriteLine("Done");
                        if (m.Logger == true)
                        {
                            File.WriteAllText("OUTPUT.LOG", Logger);
                        }
                    }
                });


            }
            catch (Exception ex)
            {
                TimeService(true);
                SetRealTime();
                Console.WriteLine(ex);
            }
        }

        private static void RecursiveSign(IEnumerable<string> cbspath)
        {

            foreach (var count in cbspath)
            {
                var countFIles = Directory.EnumerateFiles(count, "*", SearchOption.AllDirectories);
                foreach (var ext in countFIles)
                {
                    var extention = Path.GetExtension(ext);
                    if (extention == ".cab")
                    {
                        FileCount++;
                    }
                }
            }
            foreach (var dir in cbspath)
            {
                var enumFiles = Directory.EnumerateFiles(dir);

                foreach (var files in enumFiles)
                {
                    var extention = Path.GetExtension(files);
                    if (extention == ".cab")
                    {
                        Console.WriteLine(Path.GetFileName(files));
                        Console.Title = ($"{Path.GetFileName(files)} - [{CurrentCount}/{FileCount}]");
                        Process signcmd = new Process();
                        signcmd.StartInfo.FileName = "signtool.exe";
                        signcmd.StartInfo.Arguments = $"sign /v /s my /i \"Windows Phone Intermediate 2013\" /n \"Windows Phone OEM Test Cert 2013 (TEST ONLY)\" /fd SHA256 \"{files}\"";
                        signcmd.StartInfo.RedirectStandardOutput = true;
                        signcmd.StartInfo.RedirectStandardError = true;
                        signcmd.StartInfo.RedirectStandardInput = true;
                        signcmd.StartInfo.UseShellExecute = false;
                        signcmd.StartInfo.CreateNoWindow = true;
                        Console.WriteLine(signcmd.StartInfo.FileName + " " + signcmd.StartInfo.Arguments);
                        signcmd.Start();
                        signcmd.OutputDataReceived += Signcmd_OutputDataReceived;
                        signcmd.ErrorDataReceived += Signcmd_ErrorDataReceived;
                        signcmd.BeginOutputReadLine();
                        signcmd.BeginErrorReadLine();
                        signcmd.WaitForExit();

                        CurrentCount++;
                        Console.WriteLine(Environment.NewLine);

                    }
                }
            }
        }

        private static void SingularSign(string cbspath)
        {
            var enumFiles = Directory.EnumerateFiles(cbspath);
            foreach (var count in enumFiles)
            {
                var extention = Path.GetExtension(count);
                if (extention == ".cab")
                {
                    FileCount++;
                }
            }
            Console.WriteLine("[Signing Packages]");
            foreach (var files in enumFiles)
            {
                var extention = Path.GetExtension(files);
                if (extention == ".cab")
                {
                    Console.WriteLine(Path.GetFileName(files));
                    Console.Title = ($"{Path.GetFileName(files)} - [{CurrentCount}/{FileCount}]");
                    Process signcmd = new Process();
                    signcmd.StartInfo.FileName = "signtool.exe";
                    signcmd.StartInfo.Arguments = $"sign /v /s my /i \"Windows Phone Intermediate 2013\" /n \"Windows Phone OEM Test Cert 2013 (TEST ONLY)\" /fd SHA256 \"{files}\"";
                    signcmd.StartInfo.RedirectStandardOutput = true;
                    signcmd.StartInfo.RedirectStandardError = true;
                    signcmd.StartInfo.RedirectStandardInput = true;
                    signcmd.StartInfo.UseShellExecute = false;
                    signcmd.StartInfo.CreateNoWindow = true;
                    Console.WriteLine(signcmd.StartInfo.FileName + " " + signcmd.StartInfo.Arguments);
                    signcmd.Start();
                    signcmd.OutputDataReceived += Signcmd_OutputDataReceived;
                    signcmd.ErrorDataReceived += Signcmd_ErrorDataReceived;
                    signcmd.BeginOutputReadLine();
                    signcmd.BeginErrorReadLine();
                    signcmd.WaitForExit();

                    CurrentCount++;
                    Console.WriteLine(Environment.NewLine);

                }
            }
        }

        private static void Signcmd_ErrorDataReceived(object sender, DataReceivedEventArgs e)
        {
            if (!String.IsNullOrEmpty(e.Data))
            {
                Console.WriteLine(e.Data);
                Logger += $"{e.Data}\n";
                if (e.Data.Contains("Number of errors: 1"))
                {
                    ErrorCount++;
                }
            }
        }

        private static void Signcmd_OutputDataReceived(object sender, DataReceivedEventArgs e)
        {
            if (!String.IsNullOrEmpty(e.Data))
            {
                Console.WriteLine(e.Data);
                Logger += $"{e.Data}\n";
                if (e.Data.Contains("Number of errors: 1"))
                {
                    ErrorCount++;
                }
            }
        }

        public static void SetRealTime()
        {
            
            //var currentTime = GetNetworkTime();

            Console.WriteLine("Setting Time and Date back to: " + savedDate.Date.ToString() + Environment.NewLine);

            SYSTEMTIME st = new SYSTEMTIME();
            // All of these must be short
            st.wYear = (short)savedDate.Year;
            st.wMonth = (short)savedDate.Month;
            st.wDay = (short)savedDate.Day;
            st.wHour = (short)DateTime.Now.Hour;
            st.wMinute = (short)DateTime.Now.Minute;
            st.wSecond = (short)DateTime.Now.Second;

            // invoke the SetSystemTime method now
            SetSystemTime(ref st);

        }


        public static DateTime GetNetworkTime()
        {
            //default Windows time server
            const string ntpServer = "time.windows.com";

            // NTP message size - 16 bytes of the digest (RFC 2030)
            var ntpData = new byte[48];

            //Setting the Leap Indicator, Version Number and Mode values
            ntpData[0] = 0x1B; //LI = 0 (no warning), VN = 3 (IPv4 only), Mode = 3 (Client Mode)

            var addresses = Dns.GetHostEntry(ntpServer).AddressList;

            //The UDP port number assigned to NTP is 123
            var ipEndPoint = new IPEndPoint(addresses[0], 123);
            //NTP uses UDP

            using (var socket = new Socket(AddressFamily.InterNetwork, SocketType.Dgram, ProtocolType.Udp))
            {
                socket.Connect(ipEndPoint);

                //Stops code hang if NTP is blocked
                socket.ReceiveTimeout = 3000;

                socket.Send(ntpData);
                socket.Receive(ntpData);
                socket.Close();
            }

            //Offset to get to the "Transmit Timestamp" field (time at which the reply 
            //departed the server for the client, in 64-bit timestamp format."
            const byte serverReplyTime = 40;

            //Get the seconds part
            ulong intPart = BitConverter.ToUInt32(ntpData, serverReplyTime);

            //Get the seconds fraction
            ulong fractPart = BitConverter.ToUInt32(ntpData, serverReplyTime + 4);

            //Convert From big-endian to little-endian
            intPart = SwapEndianness(intPart);
            fractPart = SwapEndianness(fractPart);

            var milliseconds = (intPart * 1000) + ((fractPart * 1000) / 0x100000000L);

            //**UTC** time
            var networkDateTime = (new DateTime(1900, 1, 1, 0, 0, 0, DateTimeKind.Utc)).AddMilliseconds((long)milliseconds);

            return networkDateTime.ToLocalTime();
        }

        // stackoverflow.com/a/3294698/162671
        static uint SwapEndianness(ulong x)
        {
            return (uint)(((x & 0x000000ff) << 24) +
                           ((x & 0x0000ff00) << 8) +
                           ((x & 0x00ff0000) >> 8) +
                           ((x & 0xff000000) >> 24));
        }

        private static void TimeService(bool start)
        {
            Process svc = new Process();
            svc.StartInfo.FileName = "cmd.exe";
            svc.StartInfo.UseShellExecute = false;
            svc.StartInfo.CreateNoWindow = true;
            switch (start)
            {
                case true:
                    svc.StartInfo.Arguments = "/c sc config \"W32Time\" start= auto && net start W32Time";
                    break;
                case false:
                    svc.StartInfo.Arguments = "/c sc config \"W32Time\" start= disabled && net stop W32Time";
                    break;
            }
            svc.Start();
            svc.WaitForExit();
        }
    }


}

using ArduinoKnob.Properties;
using System;
using System.Windows.Forms;
using System.IO.Ports;
using System.Runtime.InteropServices;
using System.Management;
using System.Collections.Generic;
using System.Threading;

namespace ArduinoKnob
{
    static class Program
    {
        /// <summary>
        /// Hlavní vstupní bod aplikace.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MyCustomApplicationContext());
        }
    }


    public class MyCustomApplicationContext : ApplicationContext
    {
        private NotifyIcon trayIcon;
        static SerialPort ser;
        [DllImport("user32.dll")]
        static extern void keybd_event(byte bVk, byte bScan, uint dwFlags,
        UIntPtr dwExtraInfo);
        public const int KEYEVENTF_KEYUP = 0x0002;
        public const byte VK_MEDIA_PLAY_PAUSE = 0xB3;
        public const byte VK_VOLUME_UP = 0xAF;
        public const byte VK_VOLUME_DOWN = 0xAE;


        public MyCustomApplicationContext()
        {
            // Initialize Tray Icon
            trayIcon = new NotifyIcon()
            {
                Icon = Resources.Icon2,
                ContextMenu = new ContextMenu(new MenuItem[] {
                new MenuItem("Exit", Exit),
                new MenuItem("Verze:0.3")
            }),
                Visible = true
            };

            string p = null;
            do
            {
                p = FindPortByDescription("USB-SERIAL CH340");
                System.Threading.Thread.Sleep(10);
            } while (p == null);
            ser = new SerialPort(p);
            try
            {
                ser.Open();
                if (ser != null)
                {
                    trayIcon.Icon = Resources.Icon1;
                    Thread newThread = new Thread(Translate);
                    newThread.Start();
                }
            }
            catch (Exception exc)
            {
                string outp = exc.ToString();
                Console.Write(outp);
            }
        }

        static ManagementObject[] FindPorts()
        {
            try
            {
                ManagementObjectSearcher searcher = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM Win32_PnPEntity");
                List<ManagementObject> objects = new List<ManagementObject>();

                foreach (ManagementObject obj in searcher.Get())
                {
                    objects.Add(obj);
                }

                return objects.ToArray();
            }
            catch (Exception e)
            {
                Console.Write(e);
                return new ManagementObject[] { };
            }
        }

        static string[] FindAllPorts()
        {
            List<string> ports = new List<string>();

            foreach (ManagementObject obj in FindPorts())
            {
                try
                {
                    if (obj["Caption"].ToString().Contains("(COM"))
                    {
                        string comName = ParseCOMName(obj);
                        if (comName != null)
                            ports.Add(comName);
                    }
                }
                catch(Exception e)
                {
                    Console.Write(e);
                }
            }

            return ports.ToArray();
        }

        static string ParseCOMName(ManagementObject obj)         
        {
            string name = obj["Name"].ToString();
            int startIndex = name.LastIndexOf("(");
            int endIndex = name.LastIndexOf(")");


            if (startIndex != -1 && endIndex != -1)
            {
                name = name.Substring(startIndex + 1, endIndex - startIndex - 1);
                return name;
            }
            return null;
        }

        static string FindPortByDescription(string description)
        {
            foreach (ManagementObject obj in FindPorts())
            {
                try
                {
                    if (obj["Description"].ToString().ToLower().Equals(description.ToLower()))
                    {
                        string comName = ParseCOMName(obj);
                        if (comName != null)
                            return comName;
                    }
                }
                catch (Exception e)
                {
                    Console.Write(e);
                }
            }
            return null;
        }

        void Exit(object sender, EventArgs e)
        {
            trayIcon.Visible = false;
            trayIcon.Dispose();
            ser.Close();
            Environment.Exit(0);
        }

        public static void Translate()
        {
            
            while (ser != null && ser.IsOpen)
            {
                string rer = ser.ReadLine();
                string re = rer.Replace("\r", "");
                switch (re)
                {
                    case "Exit":
                        ser.Close();
                        break;
                    case "cw":
                        keybd_event(VK_VOLUME_DOWN, 0, 0, UIntPtr.Zero);
                        keybd_event(VK_VOLUME_DOWN, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
                        Console.WriteLine("cw");
                        continue;
                    case "ccw":
                        keybd_event(VK_VOLUME_UP, 0, 0, UIntPtr.Zero);
                        keybd_event(VK_VOLUME_UP, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
                        Console.WriteLine("ccw");
                        continue;
                    case "btn":
                        keybd_event(VK_MEDIA_PLAY_PAUSE, 0, 0, UIntPtr.Zero);
                        keybd_event(VK_MEDIA_PLAY_PAUSE, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
                        Console.WriteLine("btn");
                        continue;
                    default:
                        continue;
                }
            }
        }
    }
}


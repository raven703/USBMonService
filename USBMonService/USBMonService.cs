using System;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Net;
using System.Net.Http;
using Json.Net;
using System.Management;

using System.Collections.Generic;


using System.Text;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Security.Policy;

namespace USBMonService
{

  
    public partial class USBMonService : ServiceBase
    {
 
        public HttpClient httpClient = new HttpClient();


        public USBMonService()
        {
            InitializeComponent();
            eventLog2 = new EventLog();
            if (!EventLog.SourceExists("MySource2"))
            {
                EventLog.CreateEventSource("MySource2", "USBMonLog");
            }
            eventLog2.Source = "MySource2";
            eventLog2.Log = "USBMonLog";

            
        }



        protected void VolumeWatcherControl(int status)
        {
            
            
            WqlEventQuery volumeQuery = new WqlEventQuery("SELECT * FROM Win32_VolumeChangeEvent "); 
            ManagementEventWatcher volumeWatcher = new ManagementEventWatcher(volumeQuery);
            volumeWatcher.EventArrived += new EventArrivedEventHandler(VolumeEvent);
            
            if (status == 1)
            {
               volumeWatcher.Start();
                return;
            }

            if (status == 0)
            {
                volumeWatcher.Stop();
                return;
            }


        }  

        protected override void OnStart(string[] args)
        {
            eventLog2.WriteEntry("USBMon Service started.");
            // WqlEventQuery volumeQuery = new WqlEventQuery("SELECT * FROM Win32_VolumeChangeEvent "); //WHERE EventType = 2
            // ManagementEventWatcher volumeWatcher = new ManagementEventWatcher(volumeQuery);
            // volumeWatcher.EventArrived += new EventArrivedEventHandler(VolumeEvent);
            // volumeWatcher.Start();
            VolumeWatcherControl(1);
            

        }

        protected override void OnStop()
        {
            eventLog2.WriteEntry("USBMon Service stopped.");
            VolumeWatcherControl(0);

        }

        async void VolumeEvent(object sender, EventArrivedEventArgs e)
        {
            string time = DateTime.Now.ToString("dd-MM-yyyy HH:mm tt");

            int myEventType = Convert.ToInt32(e.NewEvent.Properties["EventType"].Value.ToString());
            //eventLog2.WriteEntry("event type is: " + myEventType);

            switch (myEventType)
            {
                case 2:
                    //eventLog2.WriteEntry("2: Device arrival");

                    //read file from web server with access list
                    string accessUrl = "http://usbmon.loc/access.lst/";
                    string access_list_string = await httpClient.GetStringAsync(accessUrl);
                    string[] access_list = access_list_string.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);

                    ManagementObjectSearcher theSearcher = new ManagementObjectSearcher("SELECT * FROM Win32_Volume WHERE DriveType='2'");

                    int driveCount = theSearcher.Get().Count;

                    if (driveCount > 0) //check if any removable disk present after query
                    {
                        foreach (ManagementObject currentObject in theSearcher.Get())
                        {

                            string diskSerialNumber = currentObject["SerialNumber"].ToString();
                            string diskDriveLetter = currentObject["DriveLetter"].ToString();
                            string hostName = Dns.GetHostName();
                            string diskStatus = "Device arrival";



                            if (!access_list.Contains(diskSerialNumber))
                            {


                          eventLog2.WriteEntry("Serial Number NOT in access List\n" + 
                         "Disk Letter: " + diskDriveLetter + "\n" + "Serial number: " + diskSerialNumber + "\n" +
                         "Host Name: " + hostName + "\n" + "Time: " + time);



                                string url = "http://usbmon.loc/query";

                                var myData = new DiskHttpData
                                {
                                    DiskSerialNumber = diskSerialNumber,
                                    DiskDriveLetter = diskDriveLetter,
                                    HostName = hostName,
                                    DiskStatus = diskStatus,
                                    SystemTime = time
                                };

                                var jsonHttpData = JsonNet.Serialize(myData);
                                StringContent httpContent = new StringContent(jsonHttpData);
                                HttpResponseMessage httpResult = await httpClient.PostAsync(url, httpContent);

                                //eventLog2.WriteEntry(httpResult.StatusCode.ToString());

                            }

                            else
                            {
                                eventLog2.WriteEntry("Serial Number IN access List");
                            }
                        }
                    }
                    else if (driveCount == 0)
                    {
                        eventLog2.WriteEntry("Unknown USB Device.");
                        string url = "http://usbmon.loc/query";


                        ManagementObjectSearcher theSearcher2 = new ManagementObjectSearcher("SELECT * FROM Win32_DiskDrive WHERE InterfaceType='USB'");
                        foreach (ManagementObject currentObject in theSearcher2.Get())
                        {

                            string diskSerialNumber = currentObject["SerialNumber"].ToString();
                            string diskDriveModel = currentObject["Model"].ToString();
                            string systemName = currentObject["SystemName"].ToString();
                            eventLog2.WriteEntry("Serial number: " + diskSerialNumber);
                            eventLog2.WriteEntry("Model: " + diskDriveModel);
                            eventLog2.WriteEntry("Host Name: " + systemName);
                            eventLog2.WriteEntry("Time: " + time);
                            
                            var myData = new DiskHttpData
                            {
                                DiskSerialNumber = diskSerialNumber,
                                DiskModel = diskDriveModel,
                                HostName = systemName,
                                DiskStatus = "Device arrival",
                                SystemTime = time
                            };
                            var jsonHttpData = JsonNet.Serialize(myData);
                            StringContent httpContent = new StringContent(jsonHttpData);
                            HttpResponseMessage httpResult = await httpClient.PostAsync(url, httpContent);
                        }
                    }
                    break;
                
                case 3:
                    eventLog2.WriteEntry("3: Device removal");
                   
                    break;

            }

        }



    }


    
    public class DiskHttpData
{
    public string DiskSerialNumber { get; set; }

    public string DiskModel { get; set; }
    public string DiskDriveLetter { get; set; }
    public string DiskStatus { get; set; }
    public string HostName { get; set; }
    public string SystemTime { get; set; }

}
}



//
// ================
// Shared Resources
// ================
//
// This class is a container for all shared resources that may be needed
// by the drivers served by the Local Server. 
//
// NOTES:
//
//	* ALL DECLARATIONS MUST BE STATIC HERE!! INSTANCES OF THIS CLASS MUST NEVER BE CREATED!
//
// Written by:	Bob Denny	29-May-2007
// Modified by Chris Rowland and Peter Simpson to hamdle multiple hardware devices March 2011
//
using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.Threading;

using ASCOM;
using ASCOM.Utilities;



namespace ASCOM.OpenAstroTracker
{
    /// <summary>
    /// The resources shared by all drivers and devices, in this example it's a serial port with a shared SendMessage method
    /// an idea for locking the message and handling connecting is given.
    /// In reality extensive changes will probably be needed.
    /// Multiple drivers means that several applications connect to the same hardware device, aka a hub.
    /// Multiple devices means that there are more than one instance of the hardware, such as two focusers.
    /// In this case there needs to be multiple instances of the hardware connector, each with it's own connection count.
    /// </summary>
    public static class SharedResources
    {
        // object used for locking to prevent multiple drivers accessing common code at the same time
        private static readonly object lockObject = new object();

        // Shared serial port. This will allow multiple drivers to use one single serial port.
        private static ASCOM.Utilities.Serial s_sharedSerial = new ASCOM.Utilities.Serial();		// Shared serial port
        private static int s_z = 0;     // counter for the number of connections to the serial port

        private static TraceLogger traceLogger;
        private static Profile driverProfile;
        public static string driverID = "ASCOM.OpenAstroTracker.Telescope";

        private static string comPortProfileName = "COM Port";
        private static string traceStateProfileName = "Trace Level";
        private static string latitudeProfileName = "Latitude";
        private static string longitudeProfileName = "Longitude";
        private static string elevationProfileName = "Elevation";

        private static string comPortDefault = "COM1";
        private static string traceStateDefault = "True";
        private static double latitudeDefault = 39.8283;
        private static double longitudeDefault = -98.5795;
        private static int elevationDefault = 1;

        private static string comPort;
        private static string portNum;
        public static bool traceState;
        public static double latitude;
        public static double longitude;
        public static int elevation;
        public static double PolarisRAJNow = 0.0;


        //
        // Public access to shared resources
        //

        public static TraceLogger tl
        {
            get
            {
                if (traceLogger == null)
                {
                    traceLogger = new TraceLogger("", "OpenAstroTracker.LocalServer");
                    traceLogger.Enabled = true;
                }
                return traceLogger;
            }
        }

        public static void ReadProfile()
        {

            using (Profile driverProfile = new Profile())
            {
                driverProfile.DeviceType = "Telescope";
                traceState = Convert.ToBoolean(driverProfile.GetValue(driverID, traceStateProfileName, String.Empty, traceStateDefault));
                comPort = driverProfile.GetValue(driverID, comPortProfileName, string.Empty, comPortDefault);
                latitude = Convert.ToDouble(driverProfile.GetValue(driverID, latitudeProfileName, string.Empty, latitudeDefault.ToString()));
                longitude = Convert.ToDouble(driverProfile.GetValue(driverID, longitudeProfileName, string.Empty, longitudeDefault.ToString()));
                elevation = Convert.ToInt16(driverProfile.GetValue(driverID, elevationProfileName, string.Empty, elevationDefault.ToString()));
            }
        }

        public static void WriteProfile()
        {

            using (Profile driverProfile = new Profile())
            {
                driverProfile.DeviceType = "Telescope";
                driverProfile.WriteValue(driverID, traceStateProfileName, traceState.ToString());
                driverProfile.WriteValue(driverID, comPortProfileName, comPort.ToString());
                driverProfile.WriteValue(driverID, latitudeProfileName, latitude.ToString());
                driverProfile.WriteValue(driverID, longitudeProfileName, longitude.ToString());
                driverProfile.WriteValue(driverID, elevationProfileName, elevation.ToString());
            }
        }

        #region single serial port connector
        //
        // this region shows a way that a single serial port could be connected to by multiple 
        // drivers.
        //
        // Connected is used to handle the connections to the port.
        //
        // SendMessage is a way that messages could be sent to the hardware without
        // conflicts between different drivers.
        //
        // All this is for a single connection, multiple connections would need multiple ports
        // and a way to handle connecting and disconnection from them - see the
        // multi driver handling section for ideas.
        //

        /// <summary>
        /// Shared serial port
        /// </summary>
        public static ASCOM.Utilities.Serial SharedSerial { get { return s_sharedSerial; } }

        /// <summary>
        /// number of connections to the shared serial port
        /// </summary>
        public static int connections { get { return s_z; } set { s_z = value; } }

        /// <summary>
        /// Example of a shared SendMessage method, the lock
        /// prevents different drivers tripping over one another.
        /// It needs error handling and assumes that the message will be sent unchanged
        /// and that the reply will always be terminated by a "#" character.
        /// </summary>
        /// <param name="message"></param>
        /// <returns></returns>
        public static string SendMessage(string message, Boolean reply)
        {
            lock (lockObject)
            {
                tl.LogMessage("OAT Server", "Lock Object");
                string msg = message + "#";

                if (SharedSerial.Connected && !String.IsNullOrEmpty(message))
                {
                    tl.LogMessage("Telescope", "Send message: " + msg);
                    SharedSerial.ClearBuffers();
                    SharedSerial.Transmit(msg);
                    if (reply)
                    {
                        string retVal;
                        string cmdGroup = message.Substring(1, 1);
                        switch (cmdGroup)
                        {
                            case "S":
                            case "M":
                            case "h":
                                retVal = SharedSerial.ReceiveCounted(1);
                                break;
                            default:
                                retVal = SharedSerial.ReceiveTerminated("#");
                                retVal.Replace("#", "");
                                break;
                        }
                        return retVal;
                    }
                    else return "";
                }
                else
                {
                    tl.LogMessage("OAT Server", "Not connected or Empty Message: " + message);
                    return "";
                }
            }
        }

        /// <summary>
        /// Example of handling connecting to and disconnection from the
        /// shared serial port.
        /// Needs error handling
        /// the port name etc. needs to be set up first, this could be done by the driver
        /// checking Connected and if it's false setting up the port before setting connected to true.
        /// It could also be put here.
        /// </summary>
        public static bool Connected
        {
            set
            {
                lock (lockObject)
                {
                    if (value)
                    {
                        if (s_z == 0)
                        {
                            try
                            {
                                SharedSerial.PortName = comPort;
                                SharedSerial.ReceiveTimeoutMs = 2000;
                                SharedSerial.Speed = ASCOM.Utilities.SerialSpeed.ps57600;
                                SharedSerial.Connected = true;
                                Thread.Sleep(2000);     // Wait for arduino to reset on serial connection
                                                        // Could be disabled w/ one of various "No rest on serial" hacks to arduino
                                SharedSerial.Transmit(":I#");
                            }
                            catch (System.IO.IOException exception)
                            {
                                MessageBox.Show("Serial port not opened for " + SharedResources.SharedSerial.PortName, "Invalid port state", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                SharedResources.tl.LogMessage("Serial port not opened", exception.Message);
                            }
                            catch (System.UnauthorizedAccessException exception)
                            {
                                MessageBox.Show("Access denied to serial port " + SharedResources.SharedSerial.PortName, "Access denied", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                SharedResources.tl.LogMessage("Access denied to serial port", exception.Message);
                            }
                            catch (ASCOM.DriverAccessCOMException exception)
                            {
                                MessageBox.Show("ASCOM driver exception: " + exception.Message, "ASCOM driver exception", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                            catch (System.Runtime.InteropServices.COMException exception)
                            {
                                MessageBox.Show("Serial port read timeout for port " + SharedResources.SharedSerial.PortName, "Timeout", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                SharedResources.tl.LogMessage("Serial port read timeout", exception.Message);
                            }
                            
                            s_z++;
                        }
                    }
                    else
                    {
                        s_z--;
                        if (s_z <= 0)
                        {
                            SharedSerial.Connected = false;
                        }
                    }
                }
            }
            get { return SharedSerial.Connected; }
        }

        #endregion

        #region Multi Driver handling
        // this section illustrates how multiple drivers could be handled,
        // it's for drivers where multiple connections to the hardware can be made and ensures that the
        // hardware is only disconnected from when all the connected devices have disconnected.

        // It is NOT a complete solution!  This is to give ideas of what can - or should be done.
        //
        // An alternative would be to move the hardware control here, handle connecting and disconnecting,
        // and provide the device with a suitable connection to the hardware.
        //
        /// <summary>
        /// dictionary carrying device connections.
        /// The Key is the connection number that identifies the device, it could be the COM port name,
        /// USB ID or IP Address, the Value is the DeviceHardware class
        /// </summary>
        private static Dictionary<string, DeviceHardware> connectedDevices = new Dictionary<string, DeviceHardware>();

        /// <summary>
        /// This is called in the driver Connect(true) property,
        /// it add the device id to the list of devices if it's not there and increments the device count.
        /// </summary>
        /// <param name="deviceId"></param>
        public static void Connect(string deviceId)
        {
            lock (lockObject)
            {
                if (!connectedDevices.ContainsKey(deviceId))
                    connectedDevices.Add(deviceId, new DeviceHardware());
                connectedDevices[deviceId].count++;       // increment the value
            }
        }

        public static void Disconnect(string deviceId)
        {
            lock (lockObject)
            {
                if (connectedDevices.ContainsKey(deviceId))
                {
                    connectedDevices[deviceId].count--;
                    if (connectedDevices[deviceId].count <= 0)
                        connectedDevices.Remove(deviceId);
                }
            }
        }

        public static bool IsConnected(string deviceId)
        {
            if (connectedDevices.ContainsKey(deviceId))
                return (connectedDevices[deviceId].count > 0);
            else
                return false;
        }

        #endregion


        /// <summary>
        /// Skeleton of a hardware class, all this does is hold a count of the connections,
        /// in reality extra code will be needed to handle the hardware in some way
        /// </summary>
        public class DeviceHardware
        {
            internal int count { set; get; }

            internal DeviceHardware()
            {
                count = 0;
            }
        }
        public static string SendMessageString(string message)
        {
            try
            {
                string answer = SharedResources.SendMessage(message, true);
                return answer;
            }
            catch
            {
                return "";
            }

        }

        //Add new public methods here

        public static void SendMessageBlind(string message)
        {
            try
            {
                string answer = SharedResources.SendMessage(message, false);
            }
            catch
            { }
        }
    }
}

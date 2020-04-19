using System;
using System.Threading;
using ASCOM.Astrometry.AstroUtils;
using ASCOM.DeviceInterface;
using ASCOM.Utilities;
using ASCOM.Astrometry.Transform;
using System.Collections;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace ASCOM.OpenAstroTracker {
    [Guid("be07c02f-8a5e-429f-87b1-23fe9d5f4065")]
    [ProgId("ASCOM.OpenAstroTracker.Telescope")]
    [ServedClassName("OpenAstroTracker")]
    [ClassInterface(ClassInterfaceType.None)]
    public class Telescope : ReferenceCountedObjectBase, ITelescopeV3 {
        // 
        // Driver ID and descriptive string that shows in the Chooser
        // 
        private string Version = "0.1.4.1b";
        private string _driverId;
        private static string driverDescription = "OpenAstroTracker Telescope";

        internal readonly string comPortProfileName = "COM Port"; // Constants used for Profile persistence
        internal readonly string traceStateProfileName = "Trace Level";
        internal readonly string latitudeProfileName = "Latitude";
        internal readonly string longitudeProfileName = "Longitude";
        internal readonly string elevationProfileName = "Elevation";

        internal static string comPortDefault = "COM5";
        internal static string traceStateDefault = "False";
        internal static double latitudeDefault = 39.8283;
        internal static double longitudeDefault = -98.5795;
        internal static int elevationDefault = 1;

        internal static string comPort; // Variables to hold the currrent device configuration
        internal static bool traceState;
        internal static double latitude;
        internal static double longitude;
        internal static double elevation;
        internal static double PolarisRAJNow = 0.0;
        internal static DriveRates driveRate = DriveRates.driveSidereal;
        
        public static string s_csDriverID;

        private bool connectedState; // Private variable to hold the connected state
        private Util utilities; // Private variable to hold an ASCOM Utilities object
        private AstroUtils astroUtilities; // Private variable to hold an AstroUtils object to provide the Range method

        private TraceLogger
            TL; // Private variable to hold the trace logger object (creates a diagnostic log file with information that you specify)

        private ASCOM.Utilities.Serial objSerial;
        public ASCOM.Astrometry.Transform.Transform transform;

        private bool
            connWait = false; // This is a lame hack to be able to transmit things w/ CommandString before we set connected to true.

        private bool isParked = false;
        private bool isTracking = true;
        private double targetRA;
        private double targetDec;
        private bool targetRASet = false;
        private bool targetDecSet = false;
        private Mutex mutexBlind;
        private Mutex mutexCommand;

        // 
        // Constructor - Must be public for COM registration!
        // 
        public Telescope() {
            
            _driverId = Marshal.GenerateProgIdForType(this.GetType());
            s_csDriverID = Marshal.GenerateProgIdForType(this.GetType());
            
            ReadProfile(); // Read device configuration from the ASCOM Profile store
            TL = new TraceLogger("", "OpenAstroTracker");
            TL.Enabled = traceState;
            LogMessage("Telescope", "Starting initialization");

            connectedState = false; // Initialise connected to false
            utilities = new Util(); // Initialise util object
            astroUtilities = new AstroUtils(); // Initialise new astro utiliites object

            // TODO: Implement your additional construction here
            mutexCommand = new Mutex(false, "CommMutex");
            transform = new Transform();
            transform.SetJ2000(utilities.HMSToHours("02:31:51.12"), utilities.DMSToDegrees("89:15:51.4"));
            transform.SiteElevation = SiteElevation;
            transform.SiteLatitude = SiteLatitude;
            transform.SiteLongitude = SiteLongitude;
            PolarisRAJNow = transform.RATopocentric;
            LogMessage("Telescope", "Completed initialization");
        }

        // 
        // PUBLIC COM INTERFACE ITelescopeV3 IMPLEMENTATION
        // 

        /// <summary>
        ///     ''' Displays the Setup Dialog form.
        ///     ''' If the user clicks the OK button to dismiss the form, then
        ///     ''' the new settings are saved, otherwise the old values are reloaded.
        ///     ''' THIS IS THE ONLY PLACE WHERE SHOWING USER INTERFACE IS ALLOWED!
        ///     ''' </summary>
        public void SetupDialog() {
            // consider only showing the setup dialog if not connected
            // or call a different dialog if connected
            if (IsConnected)
                System.Windows.Forms.MessageBox.Show("Already connected, just press OK");

            using (SetupDialogForm F = new SetupDialogForm()) {
                System.Windows.Forms.DialogResult result = F.ShowDialog();
                if (result == DialogResult.OK)
                    WriteProfile(); // Persist device configuration values to the ASCOM Profile store
            }
        }

        public ArrayList SupportedActions {
            get {
                ArrayList actionList = new ArrayList();
                actionList.Add("Telescope:getFirmwareVer");
                actionList.Add("Utility:JNowtoJ2000");
                LogMessage("SupportedActions Get",
                    "Returning arraylist of " + actionList.Count.ToString() + " item(s)");
                return actionList;
            }
        }

        public string Action(string ActionName, string ActionParameters) {
            if (SupportedActions.Contains(ActionName)) {
                string retVal = "255"; // Default error code
                switch (ActionName) {
                    case "Telescope:getFirmwareVer": {
                        retVal = CommandString(":GVP"); // Get firmware name
                        retVal = retVal + " " + CommandString(":GVN"); // Get firmware version number
                        break;
                    }

                    case "Utility:JNowtoJ2000": {
                        transform.SetTopocentric(System.Convert.ToDouble(ActionParameters.Split(',')[0]),
                            System.Convert.ToDouble(ActionParameters.Split(',')[1]));
                        retVal = utilities.HoursToHMS(transform.RAJ2000, ":", ":", string.Empty) + "&" +
                                 utilities.DegreesToDMS(transform.DecJ2000, "*", ":", string.Empty);
                        break;
                    }
                }

                LogMessage("Action(" + ActionName + ", " + ActionParameters + ")", retVal);
                return retVal;
            }
            else
                throw new ActionNotImplementedException("Action " + ActionName + " is not supported by this driver");
        }

        public void CommandBlind(string Command, bool Raw = false) {
            if (!connWait)
                CheckConnected("CommandBlind");
            mutexCommand.WaitOne();

            if (!Raw)
                Command = Command + "#";
            try {
                objSerial.Transmit(Command);
                LogMessage("CommandBlind", "Transmitted " + Command);
            }
            catch (Exception ex) {
                LogMessage("CommandBlind(" + Command + ")", "Error : " + ex.Message);
            }
            finally {
                mutexCommand.ReleaseMutex();
            }
        }

        public bool CommandBool(string Command, bool Raw = false) {
            // CheckConnected("CommandBool")
            // Dim ret As String = CommandString(Command, Raw)
            // TODO decode the return string and return true or false
            throw new MethodNotImplementedException("CommandBool");
        }

        public string CommandString(string Command, bool Raw = false) {
            string response;
            if (!connWait)
                CheckConnected("CommandString");
            mutexCommand.WaitOne();
            if (!Raw)
                Command = Command + "#";
            try {
                objSerial.Transmit(Command);
                LogMessage("CommandString", "Transmitted " + Command);
                string cmdGroup = Command.Substring(1, 1);
                switch (cmdGroup) {
                    case "S":
                    case "M":
                    case "h":
                    case "Q": {
                        response = objSerial.ReceiveCounted(1);
                        break;
                    }

                    default: {
                        response = objSerial.ReceiveTerminated("#");
                        response = response.Replace("#", "");
                        break;
                    }
                }

                LogMessage("CommandString", "Received " + response);
                return response;
            }
            catch (Exception ex) {
                LogMessage("CommandString(" + Command + ")", ex.Message);
                return "255";
            }
            finally {
                mutexCommand.ReleaseMutex();
            }
        }

        public bool Connected {
            get {
                LogMessage("Connected Get", IsConnected.ToString());
                return IsConnected;
            }
            set {
                if (value == IsConnected)
                    return;

                if (value) {
                    try {
                        connWait = true;
                        objSerial = new Serial {
                            PortName = comPort,
                            Speed = SerialSpeed.ps57600,
                            ReceiveTimeoutMs = 250,
                            DTREnable = false  // disables resetting on connect.
                    };
                        objSerial.Connected = true;
                        // Default of 5s is too high.  THis will have to be managed, however, for synced commands that take longer (:hP most notably)
                        
                        Thread.Sleep(2000); // Disgusting hack to work around arduino resetting when connected.
                        // I don't know of any way to poll and see if the reset has completed
                        CommandBlind(":I"); // OAT's command for entering PC Control mode
                        if (SiderealTime - PolarisRAJNow < 0)
                            CommandString(":SH" + utilities.HoursToHM(24 + (SiderealTime - PolarisRAJNow)), false);
                        else
                            CommandString(":SH" + utilities.HoursToHM(SiderealTime - PolarisRAJNow), false);
                        string sign = "+";
                        if (SiteLatitude < 0)
                            sign = "-";
                        CommandString(
                            ":SY" + sign + utilities.DegreesToDMS(90, "*", ":", string.Empty) + "." +
                            utilities.HoursToHMS(SiderealTime, ":", ":"), false);
                        LogMessage("Connected Set", "Connecting to port " + comPort);
                        connWait = false;
                        connectedState = true;
                    }
                    catch (Exception ex) {
                        throw new ASCOM.DriverException(ex.Message);
                        LogMessage("Connected Set", "Error Connecting to port " + comPort + " - " + ex.Message);
                    }
                }
                else
                    try {
                        CommandBlind(":Qq"); // OAT's command for exiting PC Control mode
                        Thread.Sleep(1000);
                        objSerial.Connected = false;
                        connectedState = false;
                        LogMessage("Connected Set", "Disconnecting from port " + comPort);
                    }
                    catch (Exception ex) {
                        LogMessage("Connected Set", "Error Disconnecting from port " + comPort + " - " + ex.Message);
                    }
            }
        }

        public string Description {
            get {
                // this pattern seems to be needed to allow a public property to return a private field
                string d = driverDescription;
                LogMessage("Description Get", d);
                return d;
            }
        }

        public string DriverInfo {
            get {
                // Dim m_version As Version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version
                // Dim s_driverInfo As String = "OpenAstroTracker ASCOM driver version: " + m_version.Major.ToString() + "." + m_version.Minor.ToString() + "." + m_version.Build.ToString()
                string s_driverInfo = "OpenAstroTracker ASCOM driver version: " + Version;
                LogMessage("DriverInfo Get", s_driverInfo);
                return s_driverInfo;
            }
        }

        public string DriverVersion {
            get {
                LogMessage("DriverVersion Get", Version);
                return Version;
            }
        }

        public short InterfaceVersion {
            get {
                LogMessage("InterfaceVersion Get", "3");
                return 3;
            }
        }

        public string Name {
            get {
                string s_name = "OAT ASCOM";
                LogMessage("Name Get", s_name);
                return s_name;
            }
        }

        public void Dispose() {
            // Clean up the tracelogger and util objects
            TL.Enabled = false;
            TL.Dispose();
            TL = null;
            utilities.Dispose();
            utilities = null;
            astroUtilities.Dispose();
            astroUtilities = null;
        }


        public void AbortSlew() {
            if (!AtPark) {
                CommandString(":Q");
                LogMessage("AbortSlew", ":Q# Sent");
            }
            else
                throw new ASCOM.ParkedException("AbortSlew");
        }

        public AlignmentModes AlignmentMode {
            get {
                LogMessage("AlignmentMode Get", "1");
                return
                    AlignmentModes
                        .algPolar; // 1 is "Polar (equatorial) mount other than German equatorial." from AlignmentModes Enumeration
            }
        }

        public double Altitude {
            get {
                LogMessage("Altitude", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("Altitude", false);
            }
        }

        public double ApertureArea {
            get {
                LogMessage("ApertureArea Get", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("ApertureArea", false);
            }
        }

        public double ApertureDiameter {
            get {
                LogMessage("ApertureDiameter Get", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("ApertureDiameter", false);
            }
        }

        public bool AtHome {
            get {
                // This property must be False if the telescope does not support homing.
                // TODO : We'll try to implement homing later.
                LogMessage("AtHome", "Get - " + false.ToString());
                return false;
            }
        }

        public bool AtPark {
            get {
                LogMessage("AtPark", "Get - " + isParked.ToString());
                return isParked; // Custom boolean we added to track parked state
            }
        }

        public IAxisRates AxisRates(TelescopeAxes Axis) {
            LogMessage("AxisRates", "Get - " + Axis.ToString());
            return new AxisRates(Axis);
        }

        public double Azimuth {
            get {
                LogMessage("Azimuth Get", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("Azimuth", false);
            }
        }

        public bool CanFindHome {
            get {
                if (!IsConnected)
                    throw new ASCOM.NotConnectedException("CanFindHome");
                LogMessage("CanFindHome", "Get - " + false.ToString());
                return false;
            }
        }

        public bool CanMoveAxis(TelescopeAxes Axis) {
            LogMessage("CanMoveAxis", "Get - " + Axis.ToString());
            switch (Axis) {
                case TelescopeAxes.axisPrimary: {
                    return false;
                }

                case TelescopeAxes.axisSecondary: {
                    return false;
                }

                case TelescopeAxes.axisTertiary: {
                    return false;
                }

                default: {
                    throw new InvalidValueException("CanMoveAxis", Axis.ToString(), "0 to 2");
                    break;
                }
            }
        }

        public bool CanPark {
            get {
                LogMessage("CanPark", "Get - " + true.ToString());
                return true;
            }
        }

        public bool CanPulseGuide {
            get {
                LogMessage("CanPulseGuide", "Get - " + true.ToString());
                return true;
            }
        }

        public bool CanSetDeclinationRate {
            get {
                LogMessage("CanSetDeclinationRate", "Get - " + false.ToString());
                return false;
            }
        }

        public bool CanSetGuideRates {
            get {
                LogMessage("CanSetGuideRates", "Get - " + false.ToString());
                return false;
            }
        }

        public bool CanSetPark {
            // ToDo  We should allow this
            get {
                LogMessage("CanSetPark", "Get - " + false.ToString());
                return false;
            }
        }

        public bool CanSetPierSide {
            get {
                LogMessage("CanSetPierSide", "Get - " + false.ToString());
                return false;
            }
        }

        public bool CanSetRightAscensionRate {
            get {
                LogMessage("CanSetRightAscensionRate", "Get - " + false.ToString());
                return false;
            }
        }

        public bool CanSetTracking {
            get {
                LogMessage("CanSetTracking", "Get - " + true.ToString());
                return true;
            }
        }

        public bool CanSlew {
            get {
                LogMessage("CanSlew", "Get - " + true.ToString());
                return true;
            }
        }

        public bool CanSlewAltAz {
            // TODO - AltAz slewing
            get {
                LogMessage("CanSlewAltAz", "Get - " + false.ToString());
                return false;
            }
        }

        public bool CanSlewAltAzAsync {
            get {
                LogMessage("CanSlewAltAzAsync", "Get - " + false.ToString());
                return false;
            }
        }

        public bool CanSlewAsync {
            get {
                LogMessage("CanSlewAsync", "Get - " + true.ToString());
                return true;
            }
        }

        public bool CanSync {
            get {
                LogMessage("CanSync", "Get - " + true.ToString());
                return true;
            }
        }

        public bool CanSyncAltAz {
            get {
                LogMessage("CanSyncAltAz", "Get - " + false.ToString());
                return false;
            }
        }

        public bool CanUnpark {
            get {
                LogMessage("CanUnpark", "Get - " + true.ToString());
                return true;
            }
        }

        public double Declination {
            get {
                double declination__1 = 0.0;
                string scopeDec = CommandString(":GD");
                LogMessage("Declination", "Get - " + scopeDec);
                declination__1 = utilities.DMSToDegrees(scopeDec);
                return declination__1;
            }
        }

        public double DeclinationRate {
            get {
                double declination = 0.0;
                LogMessage("DeclinationRate", "Get - " + declination.ToString());
                return declination;
            }
            set {
                LogMessage("DeclinationRate Set", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("DeclinationRate", true);
            }
        }

        public PierSide DestinationSideOfPier(double RightAscension, double Declination) {
            LogMessage("DestinationSideOfPier Get", "Not implemented");
            throw new ASCOM.MethodNotImplementedException("DestinationSideOfPier");
        }

        public bool DoesRefraction {
            get {
                LogMessage("DoesRefraction Get", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("DoesRefraction", false);
            }
            set {
                LogMessage("DoesRefraction Set", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("DoesRefraction", true);
            }
        }

        public EquatorialCoordinateType EquatorialSystem {
            // TODO : Determine if we're using JNow or J2000, or can use both.  Work on this.
            get {
                EquatorialCoordinateType equatorialSystem__1 = EquatorialCoordinateType.equTopocentric;
                LogMessage("DeclinationRate", "Get - " + equatorialSystem__1.ToString());
                return equatorialSystem__1;
            }
        }

        public void FindHome() {
            LogMessage("FindHome", "Not implemented");
            throw new ASCOM.MethodNotImplementedException("FindHome");
        }

        public double FocalLength {
            get {
                LogMessage("FocalLength Get", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("FocalLength", false);
            }
        }

        public double GuideRateDeclination {
            get {
                LogMessage("GuideRateDeclination Get", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("GuideRateDeclination", false);
            }
            set {
                LogMessage("GuideRateDeclination Set", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("GuideRateDeclination", true);
            }
        }

        public double GuideRateRightAscension {
            get {
                LogMessage("GuideRateRightAscension Get", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("GuideRateRightAscension", false);
            }
            set {
                LogMessage("GuideRateRightAscension Set", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("GuideRateRightAscension", true);
            }
        }

        public bool IsPulseGuiding {
            get {
                bool retVal = Convert.ToBoolean(System.Convert.ToInt32(CommandString(":GIG")));
                LogMessage("isPulseGuiding Get", retVal.ToString());
                return retVal;
            }
        }

        public void MoveAxis(TelescopeAxes Axis, double Rate) {
            // TODO This is coming
            LogMessage("MoveAxis", "Not implemented");
            throw new ASCOM.MethodNotImplementedException("MoveAxis");
        }

        public void Park() {
            if (!AtPark) {
                // LogMessage("Park", "Err : Mount already parked")
                // Throw New ASCOM.ParkedException("Already Parked")
                // Else
                CommandString(":hP", false);
                PollUntilZero(":GIS");
                isParked = true;
                LogMessage("Park", "Parked mount");
            }
        }

        public void PulseGuide(GuideDirections Direction, int Duration) {
            if (!AtPark) {
                var dirString = Enum.GetName(typeof(GuideDirections), Direction);
                var durString = Duration.ToString("0000");
                LogMessage($"PulseGuide", $"{Direction} {durString}");
                var dir = dirString.Substring(5, 1);
                CommandBlind($":MG{dir}{durString}", false);
            }
            else {
                LogMessage("PulseGuide", "Parked");
                throw new ASCOM.ParkedException("PulseGuide");
            }
        }

        public double RightAscension {
            get {
                double rightAscension__1 = 0.0;
                rightAscension__1 = utilities.HMSToHours(CommandString(":GR"));
                LogMessage("RightAscension", rightAscension__1.ToString());
                return rightAscension__1;
            }
        }

        public double RightAscensionRate {
            get {
                double rightAscensionRate__1 = 0.0;
                LogMessage("RightAscensionRate", "Get - " + rightAscensionRate__1.ToString());
                return rightAscensionRate__1;
            }
            set {
                LogMessage("RightAscensionRate Set", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("RightAscensionRate", true);
            }
        }

        public void SetPark() {
            LogMessage("SetPark", "Not implemented");
            throw new ASCOM.MethodNotImplementedException("SetPark");
        }

        public PierSide SideOfPier {
            get {
                PierSide retVal;
                if (SiderealTime < 12) {
                    if (RightAscension >= SiderealTime && RightAscension <= SiderealTime + 12)
                        retVal = PierSide.pierWest;
                    else
                        retVal = PierSide.pierEast;
                }
                else if (RightAscension <= SiderealTime && RightAscension >= SiderealTime - 12)
                    retVal = PierSide.pierEast;
                else
                    retVal = PierSide.pierWest;

                LogMessage("SideOfPier Get", Enum.GetName(typeof(PierSide), retVal));
                return retVal;
            }
            set {
                LogMessage("SideOfPier Set", "Not Implemented");
                throw new ASCOM.PropertyNotImplementedException("SideOfPier", true);
            }
        }

        public double SiderealTime {
            get {
                // now using novas 3.1
                double lst = 0.0;
                using (ASCOM.Astrometry.NOVAS.NOVAS31 novas = new ASCOM.Astrometry.NOVAS.NOVAS31()) {
                    double jd = utilities.DateUTCToJulian(DateTime.UtcNow);
                    novas.SiderealTime(jd, 0, novas.DeltaT(jd), ASCOM.Astrometry.GstType.GreenwichMeanSiderealTime,
                        ASCOM.Astrometry.Method.EquinoxBased, ASCOM.Astrometry.Accuracy.Reduced, ref lst);
                }

                // Allow for the longitude
                lst += SiteLongitude / 360.0 * 24.0;

                // Reduce to the range 0 to 24 hours
                lst = astroUtilities.ConditionRA(lst);

                LogMessage("SiderealTime", "Get - " + lst.ToString());
                return lst;
            }
        }

        public double SiteElevation {
            get {
                LogMessage("SiteElevation Get", elevation.ToString());
                return elevation;
            }
            set {
                if (value >= -300 && value <= 10000) {
                    LogMessage("SiteElevation Set", value.ToString());
                    elevation = value;
                }
                else
                    throw new ASCOM.InvalidValueException("SiteElevation");
            }
        }

        public double SiteLatitude {
            get {
                LogMessage("SiteLatitude Get", latitude.ToString());
                return latitude;
            }
            set {
                LogMessage("SiteLatitude Set", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("SiteLatitude", true);
            }
        }

        public double SiteLongitude {
            get {
                LogMessage("SiteLongitude Get", longitude.ToString());
                return longitude;
            }
            set {
                LogMessage("SiteLongitude Set", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("SiteLongitude", true);
            }
        }

        public short SlewSettleTime {
            get {
                LogMessage("SlewSettleTime Get", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("SlewSettleTime", false);
            }
            set {
                LogMessage("SlewSettleTime Set", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("SlewSettleTime", true);
            }
        }

        public void SlewToAltAz(double Azimuth, double Altitude) {
            LogMessage("SlewToAltAz", "Not implemented");
            throw new ASCOM.MethodNotImplementedException("SlewToAltAz");
        }

        public void SlewToAltAzAsync(double Azimuth, double Altitude) {
            LogMessage("SlewToAltAzAsync", "Not implemented");
            throw new ASCOM.MethodNotImplementedException("SlewToAltAzAsync");
        }

        public void SlewToCoordinates(double RightAscension, double Declination) {
            // Synchronous slew to given coordinates.  Uses PollUntilZero to wait for slew to finish
            if (RightAscension <= 24 && RightAscension >= 0 && Declination >= -90 && Declination <= 90) {
                if (!AtPark) {
                    LogMessage("SlewToCoordinates",
                        "RA " + RightAscension.ToString() + ", Dec " + Declination.ToString());
                    var strRAcmd = ":Sr" + utilities.HoursToHMS(RightAscension, ":", ":");
                    var strDeccmd = utilities.DegreesToDMS(Declination, "*", ":", "");
                    if (Declination >= 0)
                        strDeccmd = "+" + strDeccmd;
                    strDeccmd = ":Sd" + strDeccmd;
                    LogMessage("SlewToCoordinatesRACmd", strRAcmd);
                    LogMessage("SlewToCoordinatesDecCmd", strDeccmd);
                    if (CommandString(strRAcmd) == "1") {
                        if (CommandString(strDeccmd) == "1") {
                            CommandString(":MS");
                            PollUntilZero(":GIS");
                        }
                    }
                }
                else {
                    LogMessage("SlewToCoordinates", "Parked");
                    throw new ASCOM.ParkedException("SlewToCoordinates");
                }
            }
            else {
                LogMessage("SlewToCoordinates",
                    "Invalid coordinates RA: " + RightAscension.ToString() + ", Dec: " + Declination.ToString());
                throw new ASCOM.InvalidValueException("SlewToCoordinates");
            }
        }

        public void SlewToCoordinatesAsync(double RightAscension, double Declination) {
            // ASynchronous slew to coordinates.  Returns immediately after receiving response from :MS that command was accepted
            if (RightAscension <= 24 && RightAscension >= 0 && Declination >= -90 && Declination <= 90) {
                if (!AtPark) {
                    LogMessage("SlewToCoordinatesAsync",
                        "RA " + RightAscension.ToString() + ", Dec " + Declination.ToString());
                    var strRAcmd = ":Sr" + utilities.HoursToHMS(RightAscension, ":", ":");
                    var strDeccmd = utilities.DegreesToDMS(Declination, "*", ":", "");
                    if (Declination >= 0)
                        strDeccmd = "+" + strDeccmd;
                    strDeccmd = ":Sd" + strDeccmd;
                    LogMessage("SlewToCoordinatesAsyncRACmd", strRAcmd);
                    LogMessage("SlewToCoordinatesAsyncDecCmd", strDeccmd);
                    if (CommandString(strRAcmd) == "1") {
                        if (CommandString(strDeccmd) == "1")
                            CommandString(":MS");
                    }
                }
                else {
                    LogMessage("SlewToCoordinatesAsync", "Parked");
                    throw new ASCOM.ParkedException("SlewToCoordinatesAsync");
                }
            }
            else {
                LogMessage("SlewToCoordinatesAsync",
                    "Invalid coordinates RA: " + RightAscension.ToString() + ", Dec: " + Declination.ToString());
                throw new ASCOM.InvalidValueException("SlewToCoordinatesAsync");
            }
        }

        public void SlewToTarget() {
            LogMessage("SlewToTarget", TargetRightAscension.ToString() + ", " + TargetDeclination.ToString());
            SlewToCoordinates(TargetRightAscension, TargetDeclination);
        }

        public void SlewToTargetAsync() {
            LogMessage("SlewToTargetAsync", TargetRightAscension.ToString() + ", " + TargetDeclination.ToString());
            SlewToCoordinatesAsync(TargetRightAscension, TargetDeclination);
        }

        public bool Slewing {
            get {
                bool retVal = Convert.ToBoolean(System.Convert.ToInt32(CommandString(":GIS")));
                LogMessage("Slewing Get", retVal.ToString());
                return retVal;
            }
        }

        public void SyncToAltAz(double Azimuth, double Altitude) {
            LogMessage("SyncToAltAz", "Not implemented");
            throw new ASCOM.MethodNotImplementedException("SyncToAltAz");
        }

        public void SyncToCoordinates(double RightAscension, double Declination) {
            if (!AtPark) {
                if (RightAscension <= 24 && RightAscension >= 0 && Declination >= -90 && Declination <= 90) {
                    string sign = string.Empty;
                    if (Declination >= 0)
                        sign = "+";
                    string success =
                        CommandString(
                            ":SY" + sign + utilities.DegreesToDMS(Declination, "*", ":", string.Empty) + "." +
                            utilities.HoursToHMS(RightAscension, ":", ":"), false);
                    if (success == "1")
                        LogMessage("SyncToCoordinates",
                            "Synced to " + utilities.DegreesToDMS(Declination) + ", " +
                            utilities.HoursToHMS(RightAscension));
                    else {
                        LogMessage("SyncToCoordinates",
                            "Err - Failed to sync to " + utilities.DegreesToDMS(Declination) + ", " +
                            utilities.HoursToHMS(RightAscension));
                        throw new ASCOM.DriverException("SyncToCoordinates");
                    }
                }
                else {
                    LogMessage("SyncToCoordinates",
                        "Err - Invalid coordinates RA: " + RightAscension.ToString() + ", Dec: " +
                        Declination.ToString());
                    throw new ASCOM.InvalidValueException("SyncToCoordinates");
                }
            }
            else {
                LogMessage("SyncToCoordinates", "Err - Parked");
                throw new ASCOM.ParkedException("SyncToCoordinates");
            }
        }

        public void SyncToTarget() {
            if (!AtPark) {
                if (targetDecSet) {
                    if (targetRASet) {
                        string sign = string.Empty;
                        if (TargetDeclination >= 0)
                            sign = "+";
                        string success =
                            CommandString(
                                ":SY" + sign + utilities.DegreesToDMS(TargetDeclination, "*", ":", string.Empty) + "." +
                                utilities.HoursToHMS(TargetRightAscension, ":", ":"), false);
                        if (success == "1")
                            LogMessage("SyncToTarget",
                                "Synced to " + utilities.DegreesToDMS(TargetDeclination) + ", " +
                                utilities.HoursToHMS(TargetRightAscension));
                        else {
                            LogMessage("SyncToTarget",
                                "Failed to sync to " + utilities.DegreesToDMS(TargetDeclination) + ", " +
                                utilities.HoursToHMS(TargetRightAscension));
                            throw new ASCOM.DriverException("SyncToTarget");
                        }
                    }
                    else
                        throw new ASCOM.ValueNotSetException("TargetRightAscension");
                }
                else
                    throw new ASCOM.ValueNotSetException("TargetDeclination");
            }
            else {
                LogMessage("SyncToTarget", "Err - Parked");
                throw new ASCOM.ParkedException("SyncToTarget");
            }
        }

        public double TargetDeclination {
            get {
                if (targetDecSet) {
                    LogMessage("TargetDeclination Get", targetDec.ToString());
                    return targetDec;
                }
                else {
                    LogMessage("TargetDeclination Get", "Value not set");
                    throw new ASCOM.ValueNotSetException("TargetDeclination");
                }
            }
            set {
                if (value >= -90 && value <= 90) {
                    LogMessage("TargetDeclination Set", value.ToString());
                    targetDec = value;
                    targetDecSet = true;
                }
                else {
                    LogMessage("TargetDeclination Set", "Invalid Value " + value.ToString());
                    throw new ASCOM.InvalidValueException("TargetDeclination");
                }
            }
        }

        public double TargetRightAscension {
            get {
                if (targetRASet) {
                    LogMessage("TargetRightAscension Get", targetRA.ToString());
                    return targetRA;
                }
                else {
                    LogMessage("TargetRightAscension Get", "Value not set");
                    throw new ASCOM.ValueNotSetException("TargetRightAscension");
                }
            }
            set {
                if (value >= 0 && value <= 24) {
                    LogMessage("TargetRightAscension Set", value.ToString());
                    targetRA = value;
                    targetRASet = true;
                }
                else {
                    LogMessage("TargetRightAscension Set", "Invalid Value " + value.ToString());
                    throw new ASCOM.InvalidValueException("TargetRightAscension");
                }
            }
        }

        public bool Tracking {
            get {
                if (CommandString(":GIT", false) == "0") {
                    isTracking = false;
                    LogMessage("Tracking", "Get - " + false.ToString());
                }
                else {
                    isTracking = true;
                    LogMessage("Tracking", "Get - " + true.ToString());
                }

                return isTracking;
            }
            set {
                if (CommandString(":MT" + Convert.ToInt32(value).ToString(), false) == "1") {
                    isTracking = value;
                    LogMessage("Tracking Set", value.ToString());
                }
                else {
                    LogMessage("Tracking Set", "Error");
                    throw new ASCOM.DriverException("Error setting tracking state");
                }
            }
        }

        public DriveRates TrackingRate {
            get {
                LogMessage("TrackingRate Get", Enum.GetName(typeof(DriveRates), driveRate));
                return driveRate;
            }
            set {
                LogMessage("TrackingRate Set", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("TrackingRate", true);
            }
        }

        public ITrackingRates TrackingRates {
            get {
                ITrackingRates trackingRates__1 = new TrackingRates();
                LogMessage("TrackingRates", "Get - ");
                foreach (DriveRates driveRate in trackingRates__1)
                    LogMessage("TrackingRates", "Get - " + driveRate.ToString());
                return trackingRates__1;
            }
        }

        public DateTime UTCDate {
            // ToDo - Can we handle this without bothering the mount?
            get {
                DateTime utcDate__1 = DateTime.UtcNow;
                LogMessage("UTCDate", string.Format("Get - {0}", utcDate__1));
                return utcDate__1;
            }
            set { throw new ASCOM.PropertyNotImplementedException("UTCDate", true); }
        }

        public void Unpark() {
            if (AtPark) {
                // LogMessage("Unpark", "Err : Mount not parked")
                // Throw New ASCOM.DriverException("Mount not parked")
                // Else
                string unprkRet = CommandString(":hU", false);
                if (unprkRet == "1")
                    LogMessage("Unpark", "Unparked mount");
                isParked = false;
            }
        }


        // here are some useful properties and methods that can be used as required
        // to help with

#if INPROCESS

        private static void RegUnregASCOM(bool bRegister) {
            using (Profile P = new Profile() {DeviceType = "Telescope"}) {
                if (bRegister)
                    P.Register(driverID, driverDescription);
                else
                    P.Unregister(driverID);
            }
        }

        [ComRegisterFunction()]
        public static void RegisterASCOM(Type T) {
            RegUnregASCOM(true);
        }

        [ComUnregisterFunction()]
        public static void UnregisterASCOM(Type T) {
            RegUnregASCOM(false);
        }
#endif

        /// <summary>
        ///     ''' Returns true if there is a valid connection to the driver hardware
        ///     ''' </summary>
        private bool IsConnected {
            get {
                // TODO check that the driver hardware connection exists and is connected to the hardware
                return connectedState;
            }
        }

        /// <summary>
        ///     ''' Use this function to throw an exception if we aren't connected to the hardware
        ///     ''' </summary>
        ///     ''' <param name="message"></param>
        private void CheckConnected(string message) {
            if (!IsConnected)
                throw new NotConnectedException(message);
        }

        /// <summary>
        ///     ''' Read the device configuration from the ASCOM Profile store
        ///     ''' </summary>
        internal void ReadProfile() {
            using (Profile driverProfile = new Profile()) {
                driverProfile.DeviceType = "Telescope";
                traceState = Convert.ToBoolean(driverProfile.GetValue(_driverId, traceStateProfileName, string.Empty,
                    traceStateDefault));
                comPort = driverProfile.GetValue(_driverId, comPortProfileName, string.Empty, comPortDefault);
                Double.TryParse(driverProfile.GetValue(_driverId, latitudeProfileName, string.Empty, ""), out latitude);
                Double.TryParse(driverProfile.GetValue(_driverId, longitudeProfileName, string.Empty, ""),
                    out longitude);
                Double.TryParse(driverProfile.GetValue(_driverId, elevationProfileName, string.Empty, ""),
                    out elevation);
            }
        }

        /// <summary>
        ///     ''' Write the device configuration to the  ASCOM  Profile store
        ///     ''' </summary>
        internal void WriteProfile() {
            using (Profile driverProfile = new Profile()) {
                driverProfile.DeviceType = "Telescope";
                driverProfile.WriteValue(_driverId, traceStateProfileName, traceState.ToString());
                driverProfile.WriteValue(_driverId, comPortProfileName, comPort.ToString());
                driverProfile.WriteValue(_driverId, latitudeProfileName, latitude.ToString());
                driverProfile.WriteValue(_driverId, longitudeProfileName, longitude.ToString());
                driverProfile.WriteValue(_driverId, elevationProfileName, elevation.ToString());
            }
        }


        private int PollUntilZero(string command) {
            // Takes a command to be sent via CommandString, and resends every 1000ms until a 0 is returned.  Returns 0 only when complete.
            string retVal = "";
            while (retVal != "0") {
                retVal = CommandString(command, false);
                Thread.Sleep(1000);
            }

            return System.Convert.ToInt32(retVal);
        }

        private void LogMessage(string identifier, string message) {
            Debug.WriteLine($"{identifier} - {message}");
            TL.LogMessage(identifier, message);
        }
    }
}
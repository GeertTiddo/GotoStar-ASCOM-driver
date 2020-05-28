//tabs=4
// --------------------------------------------------------------------------------
// TODO fill in this information for your driver, then remove this line!
//
// ASCOM Telescope driver for GotoStar
//
// Description:	Driver developed just in order to use GotoStar with PHD2
//
// Implements:	ASCOM Telescope interface version: <To be completed by driver developer>
// Author:		Ger Dik (geert.dik@gmail.com)
//
// Edit Log:
//
// Date			Who	Vers	Description
// -----------	---	-----	-------------------------------------------------------
// dd-mmm-yyyy	XXX	6.0.0	Initial edit, created from ASCOM driver template
// --------------------------------------------------------------------------------
//


// This is used to define code in the template that is specific to one class implementation
// unused code canbe deleted and this definition removed.
#define Telescope

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;
using System.Runtime.InteropServices;

using ASCOM;
using ASCOM.Astrometry;
using ASCOM.Astrometry.AstroUtils;
using ASCOM.Utilities;
using ASCOM.DeviceInterface;
using System.Globalization;
using System.Collections;
using System.Windows.Forms;
using System.Diagnostics.Eventing.Reader;

using Serilog;

namespace ASCOM.GotoStar
{
    // 
    // Your driver's DeviceID is ASCOM.GotoStar.Telescope
    //
    // The Guid attribute sets the CLSID for ASCOM.GotoStar.Telescope
    // The ClassInterface/None addribute prevents an empty interface called
    // _GotoStar from being created and used as the [default] interface
    //
    // TODO Replace the not implemented exceptions with code to implement the function or
    // throw the appropriate ASCOM exception.
    //

    /// <summary>
    /// ASCOM Telescope Driver for GotoStar.
    /// </summary>
    [Guid("35d75f20-690a-44a1-a653-1e85d1cc1cda")]
    [ClassInterface(ClassInterfaceType.None)]
    public class Telescope : ITelescopeV3
    {
        //see https://ascom-standards.org/Help/Platform/html/T_ASCOM_DeviceInterface_ITelescopeV3.htm

        /// <summary>
        /// ASCOM DeviceID (COM ProgID) for this driver.
        /// The DeviceID is used by ASCOM applications to load the driver at runtime.
        /// </summary>
        internal static string driverID = "ASCOM.GotoStar.Telescope";
        // TODO Change the descriptive string for your driver then remove this line
        /// <summary>
        /// Driver description that displays in the ASCOM Chooser.
        /// </summary>
        private static string driverDescription = "ASCOM GotoStar driver";

        internal static string comPortProfileName = "COM Port"; // Constants used for Profile persistence
        internal static string comPortDefault = "COM1";
        internal static string traceStateProfileName = "Trace Level";
        internal static string traceStateDefault = "false";

        internal static string comPort; // Variables to hold the currrent device configuration

        /// <summary>
        /// If true debug log is active
        /// </summary>
        public static bool LogDebug { get; set; }

        /// <summary>
        /// Private variable to hold the connected state
        /// </summary>
        private bool connectedState;

        /// <summary>
        /// Private variable to hold an ASCOM Utilities object
        /// </summary>
        private Util utilities;

        /// <summary>
        /// Private variable to hold an ASCOM AstroUtilities object to provide the Range method
        /// </summary>
        private AstroUtils astroUtilities;

        /// <summary>
        /// private variable that contains the object that communicates with the mount
        /// </summary>
        private GotoStarCommunication gotoStar;
                
        /// <summary>
        /// Initializes a new instance of the <see cref="GotoStar"/> class.
        /// Must be public for COM registration.
        /// </summary>
        public Telescope()
        {
            LogDebug = false;
            
            ReadProfile(); // Read device configuration from the ASCOM Profile store
            // initialize serilog in debug or info mode            
            if (LogDebug)
            {
                Log.Logger = new LoggerConfiguration()
                .MinimumLevel.Debug()
                .WriteTo.Console()
                .WriteTo.File("logfile.log", rollingInterval: RollingInterval.Day)
                .CreateLogger();
            }
            else
            {
                Log.Logger = new LoggerConfiguration()
                .MinimumLevel.Information()
                .WriteTo.Console()
                .WriteTo.File("logfile.log", rollingInterval: RollingInterval.Day)
                .CreateLogger();
            }
            Log.Information("Starting GotoStar driver");
            Log.Debug("Telescope - Starting initialisation");
            
            connectedState = false; // Initialise connected to false
            utilities = new Util(); //Initialise util object
            astroUtilities = new AstroUtils(); // Initialise astro utilities object
            // initialize mount communication
            gotoStar = new GotoStarCommunication();
            Log.Debug("Telescope - Completed initialisation");
        }


        //
        // PUBLIC COM INTERFACE ITelescopeV3 IMPLEMENTATION
        //

        #region Common properties and methods.

        /// <summary>
        /// Displays the Setup Dialog form.
        /// If the user clicks the OK button to dismiss the form, then
        /// the new settings are saved, otherwise the old values are reloaded.
        /// THIS IS THE ONLY PLACE WHERE SHOWING USER INTERFACE IS ALLOWED!
        /// </summary>
        public void SetupDialog()
        {
            // consider only showing the setup dialog if not connected
            // or call a different dialog if connected
            if (IsConnected)
                System.Windows.Forms.MessageBox.Show("Already connected, just press OK");

            using (SetupDialogForm F = new SetupDialogForm())
            {
                var result = F.ShowDialog();
                if (result == System.Windows.Forms.DialogResult.OK)
                {
                    WriteProfile(); // Persist device configuration values to the ASCOM Profile store
                }
            }
        }

        public ArrayList SupportedActions
        {
            get
            {
                Log.Debug("SupportedActions Get", "Returning empty arraylist");
                return new ArrayList();
            }
        }

        public string Action(string actionName, string actionParameters)
        {
            Log.Debug("Action {0}, parameters {1} not implemented", actionName, actionParameters);
            throw new ASCOM.ActionNotImplementedException("Action " + actionName + " is not implemented by this driver");
        }

        public void CommandBlind(string command, bool raw)
        {
            CheckConnected("CommandBlind");
            // Call CommandString and return as soon as it finishes
            this.CommandString(command, raw);
            // or
            throw new ASCOM.MethodNotImplementedException("CommandBlind");
            // DO NOT have both these sections!  One or the other
        }

        public bool CommandBool(string command, bool raw)
        {
            CheckConnected("CommandBool");
            string ret = CommandString(command, raw);
            // TODO decode the return string and return true or false
            // or
            throw new ASCOM.MethodNotImplementedException("CommandBool");
            // DO NOT have both these sections!  One or the other
        }

        public string CommandString(string command, bool raw)
        {
            CheckConnected("CommandString");
            // it's a good idea to put all the low level communication with the device here,
            // then all communication calls this function
            // you need something to ensure that only one command is in progress at a time

            throw new ASCOM.MethodNotImplementedException("CommandString");
        }

        public void Dispose()
        {
            // Clean up the logger and util objects
            Log.Information("Terminating GotoStar driver");
            Log.CloseAndFlush();
            utilities.Dispose();
            utilities = null;
            astroUtilities.Dispose();
            astroUtilities = null;
            gotoStar?.ClosePort();
            gotoStar?.Dispose();
        }

        public bool Connected
        {
            get
            {
                Log.Information("Connected - Get {0}", IsConnected.ToString());
                return IsConnected;
            }
            set
            {
                Log.Debug("Connected - Set {0}", value.ToString());
                if (value == IsConnected)
                    return;

                if (value)
                {
                    connectedState = true;
                    Log.Information("Connected Set, Connecting to port {0}", comPort);
                    // connect GotoStar
                    if (!gotoStar.OpenPort(comPort))
                    {
                        MessageBox.Show(gotoStar.Error, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    connectedState = false;
                    Log.Information("Connected Set, Disconnecting from port {0}", comPort);
                    // disconnecting from mount
                    gotoStar.ClosePort();
                }
            }
        }

        public string Description
        {
            // TODO customise this device description
            get
            {
                Log.Debug("Description Get", driverDescription);
                return driverDescription;
            }
        }

        public string DriverInfo
        {
            get
            {
                Version version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
                // TODO customise this driver description
                string driverInfo = "Information about the driver itself. Version: " + String.Format(CultureInfo.InvariantCulture, "{0}.{1}", version.Major, version.Minor);
                Log.Debug("DriverInfo Get", driverInfo);
                return driverInfo;
            }
        }

        public string DriverVersion
        {
            get
            {
                Version version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
                string driverVersion = String.Format(CultureInfo.InvariantCulture, "{0}.{1}", version.Major, version.Minor);
                Log.Debug("DriverVersion Get", driverVersion);
                return driverVersion;
            }
        }

        public short InterfaceVersion
        {
            // set by the driver wizard
            get
            {
                Log.Information("InterfaceVersion Get - 3");
                return Convert.ToInt16("3");
            }
        }

        public string Name
        {
            get
            {
                string name = "GotoStarDriver";
                Log.Debug("Name, Get - '{0}'", name);
                return name;
            }
        }

        #endregion

        #region ITelescope Implementation
        public void AbortSlew()
        {
            if (gotoStar.AbortSlew())
            {
                Log.Debug("AbortSlew", "command executed");
            }
            else
            { 
                throw new ASCOM.DriverException("Error occurred during 'AbortSlew' command: " + gotoStar.Error);
            }
        }

        public AlignmentModes AlignmentMode
        {
            get
            {
                return AlignmentModes.algGermanPolar;             
            }
        }

        /// <summary>
        /// Supported by GotoStar
        /// </summary>
        public double Altitude
        {
            get
            {
                if (gotoStar.GetAltitude(out double altitude))
                {
                    Log.Debug("Altitude", "Get - " + utilities.DegreesToDMS(altitude, ":", ":"));
                    return altitude;
                }
                else
                {
                    throw new ASCOM.DriverException("Error occurred while retrieving altitude: " + gotoStar.Error);
                }
            }
        }

        public double ApertureArea
        {
            get
            {
                Log.Debug("ApertureArea Get", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("ApertureArea", false);
            }
        }

        public double ApertureDiameter
        {
            get
            {
                Log.Debug("ApertureDiameter Get", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("ApertureDiameter", false);
            }
        }

        public bool AtHome
        {
            get
            {
                Log.Debug("AtHome", "Get - " + false.ToString());
                return false;
            }
        }

        public bool AtPark
        {
            get
            {
                Log.Debug("AtPark", "Get - " + false.ToString());
                return false;
            }
        }

        public IAxisRates AxisRates(TelescopeAxes Axis)
        {
            Log.Debug("AxisRates", "Get - " + Axis.ToString());
            return new AxisRates(Axis);
        }

        /// <summary>
        /// Supported by GotoStar
        /// </summary>
        public double Azimuth
        {
            get
            {
                if (gotoStar.GetAzimuth(out double azimuth))
                {
                    Log.Debug("Azimuth", "Get - " + utilities.DegreesToDMS(azimuth, ":", ":"));
                    return azimuth;
                }
                else
                {
                    throw new ASCOM.DriverException("Error occurred while retrieving azimuth: " + gotoStar.Error);
                }
            }
        }

        public bool CanFindHome
        {
            get
            {
                Log.Debug("CanFindHome", "Get - " + false.ToString());
                return false;
            }
        }

        public bool CanMoveAxis(TelescopeAxes Axis)
        {
            Log.Debug("CanMoveAxis", "Get - " + Axis.ToString());
            switch (Axis)
            {
                case TelescopeAxes.axisPrimary: return false;
                case TelescopeAxes.axisSecondary: return false;
                case TelescopeAxes.axisTertiary: return false;
                default: throw new InvalidValueException("CanMoveAxis", Axis.ToString(), "0 to 2");
            }
        }

        public bool CanPark
        {
            get
            {
                Log.Debug("CanPark", "Get - " + false.ToString());
                return false;
            }
        }

        public bool CanPulseGuide
        {
            get
            {
                return true;
            }
        }

        public bool CanSetDeclinationRate
        {
            get
            {
                Log.Debug("CanSetDeclinationRate", "Get - " + false.ToString());
                return false;
            }
        }

        public bool CanSetGuideRates
        {
            get
            {
                Log.Debug("CanSetGuideRates, Get - " + true.ToString());
                return true;
            }
        }

        public bool CanSetPark
        {
            get
            {
                Log.Debug("CanSetPark, Get - " + false.ToString());
                //@^@ is te doen 
                return false;
            }
        }

        public bool CanSetPierSide
        {
            get
            {
                Log.Debug("CanSetPierSide, Get - " + false.ToString());
                return false;
            }
        }

        public bool CanSetRightAscensionRate
        {
            get
            {
                Log.Debug("CanSetRightAscensionRate", "Get - " + false.ToString());
                return false;
            }
        }

        public bool CanSetTracking
        {
            get
            {
                Log.Debug("CanSetTracking", "Get - " + false.ToString());                
                return false;
            }
        }

        public bool CanSlew
        {
            get
            {
                Log.Debug("CanSlew, Get - " + false.ToString());
                //@^@ is te doen 
                return false;
            }
        }

        public bool CanSlewAltAz
        {
            get
            {
                Log.Debug("CanSlewAltAz, Get - " + false.ToString());
                return false;
            }
        }

        public bool CanSlewAltAzAsync
        {
            get
            {
                Log.Debug("CanSlewAltAzAsync, Get - " + false.ToString());
                return false;
            }
        }

        public bool CanSlewAsync
        {
            get
            {
                Log.Debug("CanSlewAsync, Get - " + false.ToString());
                return false;
            }
        }

        public bool CanSync
        {
            get
            {
                Log.Debug("CanSync, Get - " + false.ToString());
                //@^@ is te doen 
                return false;
            }
        }

        public bool CanSyncAltAz
        {
            get
            {
                Log.Debug("CanSyncAltAz", "Get - " + false.ToString());
                return false;
            }
        }

        public bool CanUnpark
        {
            get
            {
                Log.Debug("CanUnpark", "Get - " + false.ToString());
                return false;
            }
        }

        /// <summary>
        /// Supported by GotoStar
        /// </summary>
        public double Declination
        {
            get
            {
                if (gotoStar.GetDeclination(out double declination))
                {
                    Log.Debug("Declination, Get - " + utilities.DegreesToDMS(declination, ":", ":"));
                    return declination;
                }
                else
                {
                    Log.Error("Error occurred while retrieving declination: " + gotoStar.Error);
                    throw new ASCOM.DriverException("Error occurred while retrieving declination: " + gotoStar.Error);
                }
            }
        }

        public double DeclinationRate
        {
            get
            {
                double declination = 0.0;
                Log.Debug("DeclinationRate", "Get - " + declination.ToString());
                return declination;
            }
            set
            {
                Log.Debug("DeclinationRate Set", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("DeclinationRate", true);
            }
        }

        public PierSide DestinationSideOfPier(double RightAscension, double Declination)
        {
            Log.Debug("DestinationSideOfPier Get", "Not implemented");
            throw new ASCOM.MethodNotImplementedException("DestinationSideOfPier");
        }

        /// <summary>
        /// Unsupported
        /// </summary>
        public bool DoesRefraction
        {
            get
            {
                
                Log.Debug("DoesRefraction Get", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("DoesRefraction", false);
            }
            set
            {
                Log.Debug("DoesRefraction Set", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("DoesRefraction", true);
            }
        }

        public EquatorialCoordinateType EquatorialSystem
        {
            get
            {
                EquatorialCoordinateType equatorialSystem = EquatorialCoordinateType.equTopocentric;
                Log.Debug("DeclinationRate", "Get - " + equatorialSystem.ToString());
                return equatorialSystem;
            }
        }

        public void FindHome()
        {
            Log.Debug("FindHome", "Not implemented");
            throw new ASCOM.MethodNotImplementedException("FindHome");
        }

        public double FocalLength
        {
            get
            {
                Log.Debug("FocalLength Get", "Not implemented");
                //@^@ handmatig?
                throw new ASCOM.PropertyNotImplementedException("FocalLength", false);
            }
        }

        public double GuideRateDeclination
        {
            get
            {
                if (gotoStar.GetGuideRate(out double guideRate))                
                {
                    Log.Debug("GuideRate (req. for dec) Get - {0} °/s", guideRate);
                    return guideRate;
                }
                else
                {
                    Log.Error("Error occurred while retrieving guide rate: " + gotoStar.Error);
                    throw new ASCOM.DriverException("Error occurred while retrieving guide rate: " + gotoStar.Error);
                }
            }
            set
            {
                double guideRate = value;
                if (gotoStar.SetGuideRate(guideRate))
                {
                    Log.Debug("GuideRate (req. for dec) Set - {0} °/s", guideRate);                    
                }
                else
                {
                    Log.Error("Error occurred while setting guide rate: " + gotoStar.Error);
                    throw new ASCOM.DriverException("Error occurred while setting guide rate: " + gotoStar.Error);
                }
            }
        }

        // SAME AS DECLINATION
        public double GuideRateRightAscension
        {
            get
            {
                if (gotoStar.GetGuideRate(out double guideRateRa))
                {

                    Log.Debug("GuideRate (req. for r.a.)  Get - {0} °/s", guideRateRa);
                    return guideRateRa;
                }
                else
                {
                    Log.Error("Error occurred while retrieving guide rate: " + gotoStar.Error);
                    throw new ASCOM.DriverException("Error occurred while retrieving guide rate: " + gotoStar.Error);
                }
            }
            set
            {
                double guideRateRa = value;
                if (gotoStar.SetGuideRate(guideRateRa))
                {
                    Log.Debug("GuideRate (req. for r.a.) Set - {0} °/s", guideRateRa);
                }
                else
                {
                    Log.Error("Error occurred while setting guide rate: " + gotoStar.Error);
                    throw new ASCOM.DriverException("Error occurred while setting guide rate: " + gotoStar.Error);
                }
            }
        }

        /// <summary>
        /// Supported by GotoStar
        /// </summary>
        public bool IsPulseGuiding
        {
            get
            {
                return gotoStar.BusyGuiding;
            }
        }

        public void MoveAxis(TelescopeAxes Axis, double Rate)
        {
            Log.Debug("MoveAxis", "Not implemented");
            //@^@ is te doen
            throw new ASCOM.MethodNotImplementedException("MoveAxis");
        }

        public void Park()
        {
            Log.Debug("Park", "Not implemented");
            //@^@ handmatig?
            throw new ASCOM.MethodNotImplementedException("Park");
        }

        /// <summary>
        /// Supported by GotoStar
        /// </summary>
        public void PulseGuide(GuideDirections direction, int duration)
        {
            Log.Debug("PulseGuide - To {0}, for {1} ms", direction.ToString(), duration);
            gotoStar.PulseGuide(direction, duration);
        }

        public double RightAscension
        {
            get
            {
                if (gotoStar.GetRightAscension(out double rightAscension))
                {
                    Log.Debug("RightAscension, Get - " + utilities.DegreesToDMS(rightAscension, ":", ":"));
                    return rightAscension;
                }
                else
                {
                    Log.Error("Error occurred while retrieving right ascension: " + gotoStar.Error);
                    throw new ASCOM.DriverException("Error occurred while retrieving right ascension: " + gotoStar.Error);
                }
            }
        }

        public double RightAscensionRate
        {
            get
            {
                double rightAscensionRate = 0.0;
                Log.Debug("RightAscensionRate", "Get - " + rightAscensionRate.ToString());
                return rightAscensionRate;
            }
            set
            {
                Log.Debug("RightAscensionRate Set", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("RightAscensionRate", true);
            }
        }

        public void SetPark()
        {
            Log.Debug("SetPark", "Not implemented");
            throw new ASCOM.MethodNotImplementedException("SetPark");
        }

        /// <summary>
        /// 
        /// </summary>
        public PierSide SideOfPier
        {
            get
            {
                if (gotoStar.GetSideOfPier(out PierSide pierSide))
                {
                    Log.Debug("SideOfPier - Get - {0} ", pierSide.ToString());
                    return pierSide;
                }
                else
                {
                    Log.Error("Error occurred while retrieving pier side: " + gotoStar.Error);
                    throw new ASCOM.DriverException("Error occurred while retrieving pier side: " + gotoStar.Error);
                }
            }
            set
            {
                Log.Debug("SideOfPier Set", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("SideOfPier", true);
            }
        }

        /// <summary>
        /// Supported by GotoStar
        /// </summary>
        public double SiderealTime
        {
            get
            {
                double siderealTime = 0.0;
                if (gotoStar.GetSiderealTime(out DateTime siderealTimeDT))
                {
                    siderealTime += siderealTimeDT.Hour;
                    siderealTime += siderealTimeDT.Minute / 60.0;
                    siderealTime += siderealTimeDT.Second / 3600.0;
                    siderealTime += siderealTimeDT.Millisecond / 3600000.0;
                    Log.Information("SiderealTime (from mount) - Get       {0}", siderealTime);
                    Log.Debug("             (computed for comparison {0})", CalculateSiderealTime());
                }
                else
                {
                    siderealTime = CalculateSiderealTime();
                    Log.Information("SiderealTime (computed) - Get - {0}", siderealTime);
                }                
                return siderealTime;
            }
        }

        public double SiteElevation
        {
            get
            {
                Log.Debug("SiteElevation Get, Not implemented");
                throw new ASCOM.PropertyNotImplementedException("SiteElevation", false);
            }
            set
            {
                Log.Debug("SiteElevation Set, Not implemented");
                //@^@ handmatig?
                throw new ASCOM.PropertyNotImplementedException("SiteElevation", true);
            }
        }


        /// <summary>
        /// Supported by GotoStar
        /// </summary>
        public double SiteLatitude
        {
            get
            {
                if (!gotoStar.GetSiteLatitude(out double latitude))
                {
                    throw new ASCOM.DriverException("Error occurred during 'GetSiteLatitude' " +
                        "command: " + gotoStar.Error);
                }
                return latitude;
            }
            set
            {
                if (value > 90 || value < -90)
                {
                    throw new ASCOM.InvalidValueException(string.Format("Invliad Site latitude " +
                        "{0}°; value should be between -90° and 90°.", value));
                }
                if (!gotoStar.SetSiteLatitude(value))
                {
                    throw new ASCOM.DriverException("Error occurred during 'SetSiteLatitude' " +
                        "command: " + gotoStar.Error);
                }
            }
        }

        /// <summary>
        /// Supported by GotoStar
        /// </summary>
        public double SiteLongitude
        {
            get
            {
                if (!gotoStar.GetSiteLongitude(out double longitude))
                {
                    throw new ASCOM.DriverException("Error occurred during 'GetSiteLongitude' command: " + gotoStar.Error);
                }
                return longitude;
            }
            set
            {
                if (value > 180 || value < -180)
                {
                    throw new ASCOM.InvalidValueException(string.Format("Invliad Site longitude " +
                        "{0}°; value should be between -180° and 180°.", value));
                }
                if (!gotoStar.SetSiteLongitude(value))
                {
                    throw new ASCOM.DriverException("Error occurred during 'SetSiteLongitude' command: " + gotoStar.Error);
                }
            }
        }

        public short SlewSettleTime
        {
            get
            {
                Log.Debug("SlewSettleTime Get", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("SlewSettleTime", false);
            }
            set
            {
                Log.Debug("SlewSettleTime Set", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("SlewSettleTime", true);
            }
        }

        public void SlewToAltAz(double Azimuth, double Altitude)
        {
            Log.Debug("SlewToAltAz", "Not implemented");
            throw new ASCOM.MethodNotImplementedException("SlewToAltAz");
        }

        public void SlewToAltAzAsync(double Azimuth, double Altitude)
        {
            Log.Debug("SlewToAltAzAsync", "Not implemented");
            throw new ASCOM.MethodNotImplementedException("SlewToAltAzAsync");
        }

        public void SlewToCoordinates(double RightAscension, double Declination)
        {
            Log.Debug("SlewToCoordinates", "Not implemented");
            //@^@ is te doen
            throw new ASCOM.MethodNotImplementedException("SlewToCoordinates");
        }

        public void SlewToCoordinatesAsync(double RightAscension, double Declination)
        {
            Log.Debug("SlewToCoordinatesAsync", "Not implemented");
            throw new ASCOM.MethodNotImplementedException("SlewToCoordinatesAsync");
        }

        public void SlewToTarget()
        {
            Log.Debug("SlewToTarget", "Not implemented");
            throw new ASCOM.MethodNotImplementedException("SlewToTarget");
        }

        public void SlewToTargetAsync()
        {
            Log.Debug("SlewToTargetAsync", "Not implemented");
            throw new ASCOM.MethodNotImplementedException("SlewToTargetAsync");
        }


        /// <summary>
        /// Supported by GotoStar
        /// </summary>
        public bool Slewing
        {
            get
            {
                if (gotoStar.BusySlewing(out bool busySlewing))
                {
                    return busySlewing;
                }
                else
                {
                    throw new ASCOM.DriverException("An error occurred while querying slew mode: " + gotoStar.Error);
                }
            }
        }

        public void SyncToAltAz(double Azimuth, double Altitude)
        {
            Log.Debug("SyncToAltAz", "Not implemented");
            throw new ASCOM.MethodNotImplementedException("SyncToAltAz");
        }

        public void SyncToCoordinates(double RightAscension, double Declination)
        {
            Log.Debug("SyncToCoordinates", "Not implemented");
            throw new ASCOM.MethodNotImplementedException("SyncToCoordinates");
        }

        public void SyncToTarget()
        {
            Log.Debug("SyncToTarget", "Not implemented");
            throw new ASCOM.MethodNotImplementedException("SyncToTarget");
        }

        public double TargetDeclination
        {
            get
            {
                Log.Debug("TargetDeclination Get", "Not implemented");
                //@^@ is te doen
                throw new ASCOM.PropertyNotImplementedException("TargetDeclination", false);
            }
            set
            {
                Log.Debug("TargetDeclination Set", "Not implemented");
                //@^@ is te doen
                throw new ASCOM.PropertyNotImplementedException("TargetDeclination", true);
            }
        }

        public double TargetRightAscension
        {
            get
            {
                Log.Debug("TargetRightAscension Get", "Not implemented");
                //@^@ is te doen
                throw new ASCOM.PropertyNotImplementedException("TargetRightAscension", false);
            }
            set
            {
                Log.Debug("TargetRightAscension Set", "Not implemented");
                //@^@ is te doen
                throw new ASCOM.PropertyNotImplementedException("TargetRightAscension", true);
            }
        }

        public bool Tracking
        {
            get
            {
                if (gotoStar.GetTracking(out bool tracking))
                {
                    return tracking;
                }
                else
                {
                    throw new ASCOM.DriverException("Error occurred while retrieving tracking " +
                        "mode: " + gotoStar.Error);
                }

            }
            set
            {
                Log.Debug("Tracking Set", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("Tracking", true);
            }
        }

        /// <summary>
        /// Supported by GotoStar
        /// </summary>
        public DriveRates TrackingRate
        {
            get
            {
                if (gotoStar.GetTrackingRate(out DriveRates trackingRate))
                {
                    Log.Debug("TrackingRate", "Get - " + trackingRate.ToString());
                    return trackingRate;
                }
                else
                {
                    throw new ASCOM.DriverException("Error occurred while retrieving tracking " +
                        "rate: " + gotoStar.Error);
                }
            }
            set
            {
                if (gotoStar.SetTrackingRate(value))
                {
                    Log.Debug("TrackingRate", "Set - " + value.ToString());                    
                }
                else
                {
                    if (gotoStar.Error == "Invalid tracking rate")
                    {
                        throw new ASCOM.InvalidValueException("Invalid tracking rate '" + 
                            TrackingRate.ToString() + "'.");
                    }
                    else
                    {
                        throw new ASCOM.DriverException("Error occurred while retrieving " +
                            "tracking rate: " + gotoStar.Error);
                    }
                }
            }
        }

        /// <summary>
        /// Supported by GotoStar
        /// </summary>
        public ITrackingRates TrackingRates
        {
            get
            {
                ITrackingRates trackingRates = new TrackingRates();
                Log.Debug("TrackingRates", "Get - ");                
                foreach (DriveRates driveRate in trackingRates)
                {
                    Log.Debug("TrackingRates", "Get - " + driveRate.ToString());
                }
                return trackingRates;
            }
        }

        public DateTime UTCDate
        {
            get
            {
                DateTime utcDate = DateTime.UtcNow;
                Log.Debug("TrackingRates", "Get - " + String.Format("MM/dd/yy HH:mm:ss", utcDate));
                return utcDate;
            }
            set
            {
                Log.Debug("UTCDate Set", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("UTCDate", true);
            }
        }

        public void Unpark()
        {
            Log.Debug("Unpark", "Not implemented");
            throw new ASCOM.MethodNotImplementedException("Unpark");
        }

        private double CalculateSiderealTime()
        {
            double siderealTime = double.NaN;
            // Now using NOVAS 3.1
            using (var novas = new ASCOM.Astrometry.NOVAS.NOVAS31())
            {
                var jd = utilities.DateUTCToJulian(DateTime.UtcNow);
                novas.SiderealTime(jd, 0, novas.DeltaT(jd),
                    ASCOM.Astrometry.GstType.GreenwichApparentSiderealTime,
                    ASCOM.Astrometry.Method.EquinoxBased,
                    ASCOM.Astrometry.Accuracy.Reduced, ref siderealTime);
            }
            // Allow for the longitude
            siderealTime += SiteLongitude / 360.0 * 24.0;

            // Reduce to the range 0 to 24 hours
            siderealTime = astroUtilities.ConditionRA(siderealTime);
            return siderealTime;            
        }

        #endregion

        #region Private properties and methods
        // here are some useful properties and methods that can be used as required
        // to help with driver development

        #region ASCOM Registration

        // Register or unregister driver for ASCOM. This is harmless if already
        // registered or unregistered. 
        //
        /// <summary>
        /// Register or unregister the driver with the ASCOM Platform.
        /// This is harmless if the driver is already registered/unregistered.
        /// </summary>
        /// <param name="bRegister">If <c>true</c>, registers the driver, otherwise unregisters it.</param>
        private static void RegUnregASCOM(bool bRegister)
        {
            using (var P = new ASCOM.Utilities.Profile())
            {
                P.DeviceType = "Telescope";
                if (bRegister)
                {
                    P.Register(driverID, driverDescription);
                }
                else
                {
                    P.Unregister(driverID);
                }
            }
        }

        /// <summary>
        /// This function registers the driver with the ASCOM Chooser and
        /// is called automatically whenever this class is registered for COM Interop.
        /// </summary>
        /// <param name="t">Type of the class being registered, not used.</param>
        /// <remarks>
        /// This method typically runs in two distinct situations:
        /// <list type="numbered">
        /// <item>
        /// In Visual Studio, when the project is successfully built.
        /// For this to work correctly, the option <c>Register for COM Interop</c>
        /// must be enabled in the project settings.
        /// </item>
        /// <item>During setup, when the installer registers the assembly for COM Interop.</item>
        /// </list>
        /// This technique should mean that it is never necessary to manually register a driver with ASCOM.
        /// </remarks>
        [ComRegisterFunction]
        public static void RegisterASCOM(Type t)
        {
            RegUnregASCOM(true);
        }

        /// <summary>
        /// This function unregisters the driver from the ASCOM Chooser and
        /// is called automatically whenever this class is unregistered from COM Interop.
        /// </summary>
        /// <param name="t">Type of the class being registered, not used.</param>
        /// <remarks>
        /// This method typically runs in two distinct situations:
        /// <list type="numbered">
        /// <item>
        /// In Visual Studio, when the project is cleaned or prior to rebuilding.
        /// For this to work correctly, the option <c>Register for COM Interop</c>
        /// must be enabled in the project settings.
        /// </item>
        /// <item>During uninstall, when the installer unregisters the assembly from COM Interop.</item>
        /// </list>
        /// This technique should mean that it is never necessary to manually unregister a driver from ASCOM.
        /// </remarks>
        [ComUnregisterFunction]
        public static void UnregisterASCOM(Type t)
        {
            RegUnregASCOM(false);
        }

        #endregion

        /// <summary>
        /// Returns true if there is a valid connection to the driver hardware
        /// </summary>
        private bool IsConnected
        {
            get
            {
                // TODO check that the driver hardware connection exists and is connected to the hardware
                return connectedState;
            }
        }

        /// <summary>
        /// Use this function to throw an exception if we aren't connected to the hardware
        /// </summary>
        /// <param name="message"></param>
        private void CheckConnected(string message)
        {
            if (!IsConnected)
            {
                throw new ASCOM.NotConnectedException(message);
            }
        }

        /// <summary>
        /// Read the device configuration from the ASCOM Profile store
        /// </summary>
        internal void ReadProfile()
        {
            using (Profile driverProfile = new Profile())
            {
                driverProfile.DeviceType = "Telescope";
                string s = driverProfile.GetValue(driverID, traceStateProfileName, string.Empty, traceStateDefault);
                LogDebug = Convert.ToBoolean(s);
                comPort = driverProfile.GetValue(driverID, comPortProfileName, string.Empty, comPortDefault);
            }
        }

        /// <summary>
        /// Write the device configuration to the  ASCOM  Profile store
        /// </summary>
        internal void WriteProfile()
        {
            using (Profile driverProfile = new Profile())
            {
                driverProfile.DeviceType = "Telescope";
                driverProfile.WriteValue(driverID, traceStateProfileName, LogDebug.ToString());
                if (comPort != null)
                {
                    driverProfile.WriteValue(driverID, comPortProfileName, comPort.ToString());
                }
            }
        }
        #endregion
    }
}

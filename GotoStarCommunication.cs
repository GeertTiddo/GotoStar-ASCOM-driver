using System;
using System.Diagnostics.Eventing.Reader;
using System.Globalization;
using System.IO.Ports;
using System.Threading;

using ASCOM.Astrometry.NOVASCOM;
using ASCOM.DeviceInterface;
using ASCOM.Utilities;

using Serilog;

namespace ASCOM.GotoStar
{
    public class GotoStarCommunication : IDisposable
    {
        private SerialPort comPort;
        private readonly object comportLock = new object();
        private string replyBuffer;
        private bool received;
        private double siteLatitude;
        private double siteLongitude;
        private System.Threading.Timer eastGuideTimer;
        private System.Threading.Timer northGuideTimer;
        private System.Threading.Timer westGuideTimer;
        private System.Threading.Timer southGuideTimer;
        private bool eastGuideBusy;
        private bool northGuideBusy;
        private bool westGuideBusy;
        private bool southGuideBusy;
        private TraceLogger tl;


        public bool Connected { get; set; }
        private string error;
        public string Error 
        {
            get 
            {
                // auto clear, read once
                string tmperror = error;
                error = string.Empty;
                return tmperror;
                
            }
            set 
            { 
                error = value; 
            } 
        }

        
        public bool BusyGuiding
        {
            get { return eastGuideBusy || northGuideBusy || westGuideBusy || southGuideBusy; }
        }

        // constructor
        public GotoStarCommunication()
        {
            Log.Debug("GotoStarCommunication started");
            Connected = false;
            siteLatitude = Double.NaN;
            siteLongitude = Double.NaN;

            // initialize timer variables
            eastGuideTimer = null;
            northGuideTimer = null;
            westGuideTimer = null;
            southGuideTimer = null;

            eastGuideBusy = false;
            northGuideBusy = false;
            westGuideBusy = false;
            southGuideBusy = false;
        }

        private void StartGuideTimer(System.Threading.Timer guideTimer,
            TimerCallback timerCallback, int milliSeconds)
        {
            guideTimer = new System.Threading.Timer(timerCallback, null, 0, milliSeconds);
        }

        /// <summary>
        /// Open the COM port and prepare for communication
        /// </summary>
        /// <param name="portName">"COM1" or "COM2" etc.</param>
        /// <returns></returns>
        public bool OpenPort(string portName)
        {
            bool result = false;
            replyBuffer = string.Empty;
            received = false;
            try
            {
                comPort = new SerialPort
                {
                    PortName = portName,
                    BaudRate = 9600,
                    DataBits = 8,
                    StopBits = StopBits.One,
                    Handshake = Handshake.None,
                    Parity = Parity.None
                };
                comPort.DataReceived += OnDataReceived;
                comPort.Open();
                result = true;
            }
            catch (Exception ex)
            {
                Error = "Exception while opening COM port: " + ex.Message;
                Log.Error("GotoStarCommunication: Exception while opening COM port: " + ex.Message);
            }
            return result;
        }

        /// <summary>
        /// Close the com port
        /// </summary>
        public void ClosePort()
        {
            AbortGuiding();
            if (comPort != null)
            {
                comPort.DataReceived -= OnDataReceived;
                comPort.Close();
            }
        }

        /// <summary>
        /// Test the connection to the GotoStar controller bij requesting the command language version
        /// and the controller version
        /// </summary>
        /// <returns>the controller version or 'Error' if something went awry</returns>
        public string TestConnection()
        {
            // request command language version
            if (RequestReply(":V#", out string languageVersion))
            {
                // request controller function
                if (RequestReply(":Vs#", out string controllerVersion))
                {
                    return controllerVersion;
                }
            }
            return "Error";
        }

        /// <summary>
        /// Returns the site latitude. Note that the ":Gt#" command retrieves the site latitude 
        /// as it was manually entered on the controller. Values entered via RS-232 are ignored 
        /// by this command. Therefore, if the latitude was previously set by RS-232 a stored 
        /// value is returned.
        /// </summary>
        /// <param name="latitude"></param>
        /// <returns>true for success</returns>
        public bool GetSiteLatitude(out double latitude)
        {
            bool success = false;
            latitude = double.NaN;
            // if a previously set latitude value is available, use it
            if (!double.IsNaN(siteLatitude))
            {
                latitude = this.siteLatitude;
                success = true;
            }
            // no previously set value available, retrieve latitude from GotoStar
            else
            {
                if (RequestReply(":Gt#", out string latitudeString))
                {
                    latitude = degStr2Num(latitudeString);
                    success = true;
                }
            }
            return success;
        }

        /// <summary>
        /// Sets the latitude of the observing site. This value sent by the ":St sDD*MM:SS#
        /// "command is not persistent through a power cycle of GotosStar. Yet it is taken 
        /// into account by GotoStar's calculations of the mount's position.
        /// </summary>
        /// <param name="latitude"></param>
        /// <returns></returns>
        public bool SetSiteLatitude(double latitude)
        {
            bool success = false;
            if (RequestReply(string.Format(":St {0}#", num2DegStr2(latitude)), out string setLatitudeConfirm))
            {
                if (setLatitudeConfirm == "1")
                {
                    // store the site latitude, because GotoStar won't remember!
                    this.siteLatitude = latitude;
                    success = true;
                }
            }
            return success;
        }

        /// <summary>
        /// Returns the site longitude. Note that the ":GG#" command retrieves the site 
        /// longitude as it was manually entered on the controller. Values entered via 
        /// RS-232 are ignored by this command. Therefore, if the longitude was previously set 
        /// by RS-232, a stored value is returned.
        /// </summary>
        /// <param name="latitude"></param>
        /// <returns>true for success</returns>
        public bool GetSiteLongitude(out double longitude)
        {
            bool success = false;
            longitude = double.NaN;
            // if a previously set longitude value is available, use it
            if (!double.IsNaN(siteLongitude))
            {
                longitude = this.siteLongitude;
                success = true;
            }
            // no previously set value available, retrieve longitude from GotoStar
            else
            {
                if (RequestReply(":Gg#", out string longitudeString))
                {
                    longitude = degStr2Num(longitudeString);
                    success = true;
                }
            }
            return success;
        }

        /// <summary>
        /// Sets the longitude of the observing site. This value sent by the ":Sg sDDD*MM:SS#" 
        /// command is not persistent through a power cycle of GotosStar. Yet it is taken into
        /// account by GotoStar's calculations of the mount's position.
        /// </summary>
        /// <param name="longitude"></param>
        /// <returns></returns>
        public bool SetSiteLongitude(double longitude)
        {
            bool success = false;
            if (RequestReply(string.Format(":Sg {0}#", num2DegStr2(longitude, true)), out string setLongitudeConfirm))
            {
                if (setLongitudeConfirm == "1")
                {
                    // store the site longitude, because GotoStar won't remember!
                    this.siteLongitude = longitude;
                    success = true;
                }
            }
            return success;
        }

        public bool GetSiderealTime(out DateTime siderealTime)
        {
            bool success = false;
            siderealTime = DateTime.MinValue;
            // retrieve the mount's sidereal time
            if (RequestReply(":GS#", out string siderealTimeString))
            {
                siderealTime = DateTime.Parse(siderealTimeString.Substring(0,10));
                success = true;
            }
            return success;
        }

        public bool GetLocalTime(out DateTime localTime)
        {
            bool success = false;
            localTime = DateTime.MinValue;
            // retrieve the mount's local time
            if (RequestReply(":GL#", out string localTimeString))
            {
                localTime = DateTime.Parse(localTimeString.Substring(0, 10));
                success = true;
            }
            return success;
        }

        /// <summary>
        /// Format: 
        /// E HH:MM#
        /// W HH:MM#
        /// </summary>
        /// <param name="UTCoffset"></param>
        /// <returns></returns>
        public bool GetSiteUTCoffset(out int UTCoffset)
        {
            bool success = false;
            UTCoffset = int.MinValue;
            // retrieve the UTC offset 
            if (RequestReply(":GG#", out string UTCoffsetString))
            {
                UTCoffset = int.Parse(UTCoffsetString.Substring(1, 2));
                if (UTCoffsetString[0] == 'W')
                {
                    UTCoffset *= -1;
                }
                success = true;
            }
            return success;
        }

        public bool SetSiteUTCoffset(int UTCoffset)
        {
            bool success = false;
            if (RequestReply(string.Format(":SG {0}{1:d2}#", UTCoffset > 0 ? '+' : '-', Math.Abs(UTCoffset)), out string setUtcOffsetConfirm))
            {
                success = setUtcOffsetConfirm == "1";
            }
            return success;
        }

        #region retrieve positions

        /// <summary>
        /// Retrieve the mount's current altitude
        /// </summary>
        /// <param name="altitude"></param>
        /// <returns>true for success</returns>
        public bool GetAltitude(out double altitude)
        {
            bool success = false;
            altitude = double.NaN;
            if (RequestReply(":GA#", out string altitudeString))
            {
                altitude = degStr2Num(altitudeString);
                success = true;
            }
            return success;
        }

        /// <summary>
        /// Retrieve the mount's current azimuth
        /// </summary>
        /// <param name="azimuth"></param>
        /// <returns>true for success</returns>
        public bool GetAzimuth(out double azimuth)
        {
            bool success = false;
            azimuth = double.NaN;
            if (RequestReply(":GZ#", out string azimuthString))
            {
                azimuth = degStr2Num(azimuthString);
                success = true;
            }
            return success;
        }

        /// <summary>
        /// Retrieve the mount's current rightAscension
        /// </summary>
        /// <param name="rightAscension"></param>
        /// <returns>true for success</returns>
        public bool GetRightAscension(out double rightAscension)
        {
            bool success = false;
            rightAscension = double.NaN;
            if (RequestReply(":GR#", out string rightAscensionString))
            {
                rightAscension = TimeStr2Num(rightAscensionString);
                success = true;
            }
            return success;
        }

        /// <summary>
        /// Retrieve the mount's current declination
        /// </summary>
        /// <param name="declination"></param>
        /// <returns>true for success</returns>
        public bool GetDeclination(out double declination)
        {
            bool success = false;
            declination = double.NaN;
            if (RequestReply(":GD#", out string declinationString))
            {
                declination = degStr2Num(declinationString);
                success = true;
            }
            return success;
        }

        public bool GetSideOfPier(out PierSide pierSide)
        {
            bool success = false;
            pierSide = PierSide.pierUnknown;
            if (RequestReply(":pS#", out string pierSideReply))
            {
                switch(pierSideReply)
                {
                    case "East#":
                        pierSide = PierSide.pierEast;
                        success = true;
                        break;
                    case "West#":
                        pierSide = PierSide.pierWest;
                        success = true;
                        break;
                }                
            }
            return success;
        }

        #endregion retrieve positions

        #region move commands

        /// <summary>
        /// Moves the mount in guiding speed to a specified direction for a specified duration
        /// If the mount is not slewing, a move command is issued. Immediately a timer is started 
        /// for the specified duration, which will stop the moving if it expires.  
        /// </summary>
        /// <param name="direction"></param>
        /// <param name="duration"></param>
        /// <returns>true if succesful</returns>
        public bool PulseGuide(GuideDirections direction, int duration)
        {
            bool success = false;
            if (BusySlewing(out bool busySlewing))
            {
                if (!busySlewing)
                {
                    // go into guiding mode. Issue this command always, because
                    // slewing might have been selected manually on the controller
                    if (!Request(":RG#"))
                    {
                        return success;
                    }
                    char directionChar = ' ';
                    switch (direction)
                    {
                        case GuideDirections.guideEast:
                            if (!eastGuideBusy)
                            {
                                directionChar = 'e';
                            }
                            break;
                        case GuideDirections.guideNorth:
                            if (!northGuideBusy)
                            {
                                directionChar = 'n';
                            }
                            break;
                        case GuideDirections.guideSouth:
                            if (!southGuideBusy)
                            {
                                directionChar = 's';
                            }
                            break;
                        case GuideDirections.guideWest:
                            if (!westGuideBusy)
                            {
                                directionChar = 'w';
                            }
                            break;
                    }
                    if (directionChar != ' ')
                    {
                        if (Request(string.Format(":M{0}#", directionChar)))
                        {
                            switch (direction)
                            {
                                case GuideDirections.guideEast:
                                    eastGuideBusy = true;
                                    eastGuideTimer =
                                        new System.Threading.Timer(new TimerCallback(OnEastGuideTimer), null, duration, duration);
                                    Log.Debug("- Started guiding East. ({0} ms)", duration);
                                    success = true;
                                    break;
                                case GuideDirections.guideNorth:
                                    northGuideBusy = true;
                                    northGuideTimer =
                                        new System.Threading.Timer(new TimerCallback(OnNorthGuideTimer), null, duration, duration);
                                    Log.Debug("- Started guiding North. ({0} ms)", duration);
                                    success = true;
                                    break;
                                case GuideDirections.guideSouth:
                                    southGuideBusy = true;
                                    southGuideTimer =
                                        new System.Threading.Timer(new TimerCallback(OnSouthGuideTimer), null, duration, duration);
                                    Log.Debug("- Started guiding South. ({0} ms)", duration);
                                    success = true;
                                    break;
                                case GuideDirections.guideWest:
                                    westGuideBusy = true;
                                    westGuideTimer =
                                        new System.Threading.Timer(new TimerCallback(OnWestGuideTimer), null, duration, duration);
                                    Log.Debug("- Started guiding West. ({0} ms)", duration);
                                    success = true;
                                    break;
                            }// switch
                        }// move request succesful
                    }// direction valid
                }// mount is not slewing 
            }// able to determine slewing
            return success;
        }// PulseGuide

        /// <summary>
        /// If the mount is slewing, then stop motion in all axes
        /// </summary>
        /// <returns>true if succesful</returns>
        public bool AbortSlew()
        {
            bool success = false;
            // is the mount slewing?
            if (BusySlewing(out bool busySlewing))
            {
                if (busySlewing)
                {
                    // stop motion in all axes
                    if (Request(":Q#"))
                    {
                        success = true;
                    }
                }
                else
                {
                    success = true;
                }
            }
            return success;
        }// AbortSlew

        public void AbortGuiding()
        {
            if (eastGuideBusy)
            {
                OnEastGuideTimer(null);
            }
            if (northGuideBusy)
            {
                OnNorthGuideTimer(null);
            }
            if (westGuideBusy)
            {
                OnWestGuideTimer(null);
            }
            if (southGuideBusy)
            {
                OnSouthGuideTimer(null);
            }
        }

        /// <summary>
        /// Determines whther the mount is slewing or not
        /// </summary>
        /// <param name="busySlewing">true if slewing</param>
        /// <returns>true if success</returns>
        public bool BusySlewing(out bool busySlewing)
        {
            busySlewing = false;
            bool success = false;
            if (RequestReply(":SE?#", out string slewmode))
            {
                switch (slewmode)
                {
                    case "0":
                        busySlewing = false;
                        success = true;
                        break;
                    case "1":
                        busySlewing = true;
                        success = true;
                        break;
                    default:
                        // unexpected response
                        success = false;
                        break;
                }
            }
            return success;
        }// BusySlewing

        /// <summary>
        /// not functional
        /// </summary>
        /// <param name="on"></param>
        /// <returns></returns>
        public bool SetTracking(bool on)
        {
            bool success = false;

            if (Request(on ? ":STON#" : ":STOFF#"))
            {
                success = true;
            }
            return success;
        }

        /// <summary>
        /// Determine if the mount is tracking by the change in the RA value measured over a
        /// short time interval.
        /// If the mount is slewing or guiding during the execution of this method, the out-
        /// come will be false, although the mount could return to tracking when the slew or 
        /// guide action has finished.
        /// </summary>
        /// <param name="tracking">true if the mount is tracking</param>
        /// <returns>true for success</returns>
        public bool GetTracking(out bool tracking)
        {
            bool success = false;
            tracking = false;
            // Get the first RA sample
            if (GetRightAscension(out double rightAscension1))
            {
                // wait
                Thread.Sleep(1200);
                // get the second RA sample
                if (GetRightAscension(out double rightAscension2))
                {
                    // determine the difference in minutes
                    double raChangeMin = Math.Abs((rightAscension2 - rightAscension1) * 60);
                    // if the difference is small enough, assume tracking is active
                    tracking = raChangeMin < 0.0000001;
                    success = true;
                }
            }
            return success;
        }// GetTracking

        #endregion move commands

        #region Rates and Speeds

        public bool GetTrackingRate(out DriveRates trackingRate)
        {
            bool success = false;
            trackingRate = DriveRates.driveSidereal;
            if (RequestReply(":GTR#", out string trackingRateReply))
            {
                switch(trackingRateReply)
                {
                    case "0":
                        trackingRate = DriveRates.driveSidereal;
                        success = true;
                        break;
                    case "1":
                        trackingRate = DriveRates.driveSolar;
                        success = true; 
                        break;
                    case "2":
                        trackingRate = DriveRates.driveLunar;
                        success = true; 
                        break;
                }// switch
            }
            return success;
        }// GetTrackingRate

        /// <summary>
        /// Sets the tracking rate and also starts tracking
        /// </summary>
        /// <param name="trackingRate"></param>
        /// <returns></returns>
        public bool SetTrackingRate(DriveRates trackingRate)
        {
            bool success = false;
            char trackingChar = ' ';
            switch (trackingRate)
            {
                case DriveRates.driveSidereal:
                    trackingChar = '0';
                    break;
                case DriveRates.driveSolar:
                    trackingChar = '1'; 
                    break;
                case DriveRates.driveLunar:
                    trackingChar = '2'; 
                    break;
            }// switch            
            if (trackingChar != ' ')
            {
                if (RequestReply(string.Format(":STR{0}#",trackingChar), out string setTrackingRateReply))
                {
                    success = setTrackingRateReply == "1";
                }
            }
            else
            {
                this.Error = "Invalid tracking rate";
            }
            return success;
        }// SetTrackingRate

        /// <summary>
        /// obtain current guide rate from mount
        /// </summary>
        /// <param name="guideRateFactor">guide rate as expressed in degrees / second</param>
        /// <returns></returns>
        public bool GetGuideRate(out double guideRate)
        {
            bool success = false;
            guideRate = double.NaN;
            double guideRateFactor = 0;
            if (RequestReply(":GGS#", out string guideRateReply))
            {
                switch (guideRateReply)
                {
                    case "0":
                        guideRateFactor = 1.0;
                        success = true;
                        break;
                    case "1":
                        guideRateFactor = 0.8;
                        success = true;
                        break;
                    case "2":
                        guideRateFactor = 0.6; 
                        success = true;
                        break;
                    case "3":
                        guideRateFactor = 0.4;
                        success = true;
                        break;
                }// switch
                if (success)
                {
                    // convert factor to degrees / second
                    guideRate = guideRateFactor * (360.0 / (24.0 * 60.0 * 60.0));
                }
            }
            return success;
        }// GetGuideRate

        /// <summary>
        /// Sets the guide rate. Select the rate closest to the requested value
        /// IMPORTANT: this command messes up the site's UTC offset. Therefor 
        /// this is stored beforehand and corrected afterwards.
        /// </summary>
        /// <param name="guideRate">guide rate as expressed in degrees / second</param>
        /// <returns></returns>
        public bool SetGuideRate(double guideRate)
        {
            bool success = false;
            // preserve site's UTC offset
            if (!GetSiteUTCoffset(out int utcOffset))
            {
                return success;
            }
            double guideRateFactor = guideRate / ((360.0 / (24.0 * 60.0 * 60.0)));
            char rateChar = ' ';
            // 
            if (guideRateFactor > 0.9)
            {
                // assume factor 1.0
                rateChar = '0';
            }
            else if (guideRateFactor <= 0.9 && guideRateFactor > 0.7)
            {
                // assume factor 0.8
                rateChar = '1';
            }
            else if (guideRateFactor <= 0.7 && guideRateFactor > 0.5)
            {
                // assume factor 0.6
                rateChar = '2';
            }
            else if (guideRateFactor <= 0.5)
            {
                // assume factor 0.4
                rateChar = '3';
            }            
            if (Request(string.Format(":SGS{0}#", rateChar)))
            {
                Thread.Sleep(200);
                // preserve site's UTC offset
                if (SetSiteUTCoffset(utcOffset))
                {
                    success = true;
                }
            }            
            return success;
        }// SetGuideRate

        #endregion Rates and Speeds
        #region serial communication 

        /// <summary>
        /// Send a request to the mount 
        /// </summary>
        /// <param name="request"></param>
        /// <returns>true for success</returns>
        private bool Request(string request)
        {
            bool success = false;
            try
            {
                lock (this.comportLock)
                {
                    this.comPort.Write(request);
                }
                success = true;
            }
            catch (Exception ex)
            {
                Error = "Exception occured during serial request: " + ex.Message;
                Log.Error("GotoStarCommunication: Exception occured during serial request: " + ex.Message);
            }
            return success;
        }

        /// <summary>
        /// Send a request to the mount and wait for a reply. If no reply is received within a 
        /// second, an empty reply is returned.
        /// </summary>
        /// <param name="request"></param>
        /// <returns>true for success</returns>
        private bool RequestReply(string request, out string reply)
        {
            bool success = false;
            reply = string.Empty;
            received = false;
            replyBuffer = string.Empty;
            try
            {
                lock (this.comportLock)
                {
                    this.comPort.Write(request);
                    int timeout = 200;
                    while (!received && timeout > 0)
                    {
                        Thread.Sleep(10);
                        timeout--;
                    }
                    if (timeout > 0)
                    {
                        Connected = true;
                        reply = replyBuffer;
                        success = true;
                    }
                    else
                    {
                        Connected = false;
                        Error = "Time-out occured while waiting for reply from COM port.";
                        Log.Error("GotoStarCommunication: Time-out occured while waiting for reply from COM port.");
                        success = false;
                    }
                }
            }
            catch (Exception ex)
            {
                Error = "Exception occured during serial communication: " + ex.Message;
                Log.Error("GotoStarCommunication: Exception occured during serial communication: " + ex.Message);
            }
            return success;
        }// RequestReply

        private void OnDataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            Thread.Sleep(30);
            replyBuffer = comPort.ReadExisting();
            received = true;
        }// OnDataReceived

        #endregion serial communication

        #region Guide timer callback methods

        private void OnEastGuideTimer(Object stateInfo)
        {
            Request(":Qe#");
            Log.Debug("- Stopped guiding East.");
            eastGuideBusy = false;
            eastGuideTimer?.Dispose();
            eastGuideTimer = null;
        }

        private void OnNorthGuideTimer(Object stateInfo)
        {
            Request(":Qn#");
            Log.Debug("- Stopped guiding North.");
            northGuideBusy = false;
            northGuideTimer?.Dispose();
            northGuideTimer = null;
        }

        private void OnWestGuideTimer(Object stateInfo)
        {
            Request(":Qw#");
            Log.Debug("- Stopped guiding West.");
            westGuideBusy = false;
            westGuideTimer?.Dispose();
            westGuideTimer = null;
        }

        private void OnSouthGuideTimer(Object stateInfo)
        {
            Request(":Qs#");
            Log.Debug("- Stopped guiding South.");
            southGuideBusy = false;
            southGuideTimer?.Dispose();
            southGuideTimer = null;
        }

        #endregion Guide timer callback methods

        #region Conversion methods

        /// <summary>
        /// Converts a string witht degrees, minutes, seconds to a floating point number
        /// Formats:
        /// s[D]DD*MM:SS#, where 's' is the sign
        /// </summary>
        /// <param name=""></param>
        /// <returns></returns>
        private double degStr2Num(string dms)
        {
            double result = double.NaN;
            if (dms.Length >= 10 && dms.Contains("*") && dms.Contains(":") && dms.Contains("#"))
            {
                string sDegrees = dms.Substring(1, dms.IndexOf('*') - 1);
                string sMinutes = dms.Substring(dms.IndexOf('*') + 1, 2);
                string sSeconds = dms.Substring(dms.IndexOf(':') + 1, 2);
                result = double.Parse(sDegrees);
                result += double.Parse(sMinutes)/60;
                result += double.Parse(sSeconds)/3600;
                if (dms[0] == '-')
                { 
                    result *= -1; 
                }
            }
            return result;
        }// degStr2Num

        /// <summary>
        /// Converts a string with hours, minutes, seconds to a floating point number
        /// Formats:
        /// HH:MM:SS.S#,
        /// 0123456789
        /// </summary>
        /// <param name=""></param>
        /// <returns></returns>
        private double TimeStr2Num(string hms)
        {
            double result = double.NaN;
            if (hms.Length == 11 && hms.Contains(":") && hms.Contains("#"))
            {
                string sHours = hms.Substring(0, 2);
                string sMinutes = hms.Substring(3, 2);
                string sSeconds = hms.Substring(6, 4);
                result = double.Parse(sHours);
                result += double.Parse(sMinutes) / 60.0;
                // CultureInfo.InvariantCulture forces interpretation of decimal point
                result += double.Parse(sSeconds, CultureInfo.InvariantCulture) / 3600.0;                
            }
            return result;
        }// TimeStr2Num

        private string num2DegStr2(double degdouble, bool threeDegreePositions = false)
        {
            string result = string.Empty;
            bool negative = degdouble < 0;
            if (negative) 
            {
                degdouble *= -1;
            }
            int degrees = (int)Math.Truncate(degdouble);
            double minutes = (int)Math.Truncate((degdouble - degrees) * 60);
            int seconds = (int)Math.Round((degdouble - degrees - (minutes/60)) * 3600);
            if (threeDegreePositions)
            {
                result= string.Format("{0}{1:d3}*{2:d2}:{3:d2}", negative ? '-':' ',degrees, (int)minutes, seconds);
            }
            else
            {
                result = string.Format("{0}{1:d2}*{2:d2}:{3:d2}", negative ? '-' : ' ', degrees, (int)minutes, seconds);
            }
            return result;
        }

        #endregion Conversion methods

        public void Dispose()
        {
            eastGuideTimer?.Dispose();
            eastGuideTimer = null;
            northGuideTimer?.Dispose();
            northGuideTimer = null;
            westGuideTimer?.Dispose();
            westGuideTimer = null;
            southGuideTimer?.Dispose();
            southGuideTimer = null;
            ClosePort();
        }// Dispose

    }// class
}// namespace

//tabs=4
// --------------------------------------------------------------------------------
// TODO fill in this information for your driver, then remove this line!
//
// ASCOM Telescope driver for NSMount
//
// Description:	Lorem ipsum dolor sit amet, consetetur sadipscing elitr, sed diam 
//				nonumy eirmod tempor invidunt ut labore et dolore magna aliquyam 
//				erat, sed diam voluptua. At vero eos et accusam et justo duo 
//				dolores et ea rebum. Stet clita kasd gubergren, no sea takimata 
//				sanctus est Lorem ipsum dolor sit amet.
//
// Implements:	ASCOM Telescope interface version: <To be completed by driver developer>
// Author:		(XXX) Your N. Here <your@email.here>
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
using System.IO.Ports;

namespace ASCOM.AMount
{
    //
    // Your driver's DeviceID is ASCOM.NSMount.Telescope
    //
    // The Guid attribute sets the CLSID for ASCOM.NSMount.Telescope
    // The ClassInterface/None addribute prevents an empty interface called
    // _NSMount from being created and used as the [default] interface
    //
    // TODO Replace the not implemented exceptions with code to implement the function or
    // throw the appropriate ASCOM exception.
    //

    /// <summary>
    /// ASCOM Telescope Driver for NSMount.
    /// </summary>
    [Guid("55a35c83-da00-4b0e-b067-77b9caa83dbe")]
    [ClassInterface(ClassInterfaceType.None)]
    public class Telescope : ITelescopeV3
    {
        /// <summary>
        /// ASCOM DeviceID (COM ProgID) for this driver.
        /// The DeviceID is used by ASCOM applications to load the driver at runtime.
        /// </summary>
        internal static string driverID = "ASCOM.AMount.Telescope";
        /// <summary>
        /// Driver description that displays in the ASCOM Chooser.
        /// </summary>
        private static string driverDescription = "ASCOM AMount Driver";

        internal static string comPortProfileName = "COM Port"; // Constants used for Profile persistence
        internal static string comPortDefault = "COM1";
        internal static string traceStateProfileName = "Trace Level";
        internal static string traceStateDefault = "false";
        internal static string comPort; // Variables to hold the currrent device configuration

        /// <summary>
        /// Private variable to hold the connected state
        /// </summary>
        private bool connectedState;
        //  private bool SlewState;
        private bool MotorEnableState;

        /// <summary>
        /// Private variable to hold an ASCOM Utilities object
        /// </summary>
        private Util utilities;

        /// <summary>
        /// Private variable to hold an ASCOM AstroUtilities object to provide the Range method
        /// </summary>
        private AstroUtils astroUtilities;

        /// <summary>
        /// Variable to hold the trace logger object (creates a diagnostic log file with information that you specify)
        /// </summary>
        internal static TraceLogger tl;

        public SerialPort objSerial;

        internal bool PulseGuiding = false;

        /// <summary>
        /// Initializes a new instance of the <see cref="NSMount"/> class.
        /// Must be public for COM registration.
        /// </summary>
        public Telescope()
        {
            tl = new TraceLogger("", "NSMount");
            ReadProfile(); // Read device configuration from the ASCOM Profile store

            tl.LogMessage("Telescope", "Starting initialisation");

            connectedState = false; // Initialise connected to false
                                    //   SlewState        = false; // Initialise Slewing false
            MotorEnableState = true;  // Initialise Motor enable

            utilities = new Util(); //Initialise util object
            astroUtilities = new AstroUtils(); // Initialise astro utilities object
            //TODO: Implement your additional construction here

            tl.LogMessage("Telescope", "Completed initialisation");
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
                tl.LogMessage("SupportedActions Get", "Returning empty arraylist");
                return new ArrayList();
            }
        }

        public string Action(string actionName, string actionParameters)
        {
            LogMessage("", "Action {0}, parameters {1} not implemented", actionName, actionParameters);
            throw new ASCOM.ActionNotImplementedException("Action " + actionName + " is not implemented by this driver");
        }

        public void CommandBlind(string command, bool raw)
        {
            CheckConnected(command);
            AnswerString(false);

            // Call CommandString and return as soon as it finishes

            //or: this.CommandString(command, raw); // DO NOT have both these sections!  One or the other

            throw new ASCOM.MethodNotImplementedException("CommandBlind");

        }

        public bool CommandBool(string command, bool raw)
        {
            CheckConnected(command);
            string ret = AnswerString(raw); //ѕараметр raw указывает, нужен ли завершающий # в строке ответа

            // TODO decode the return string and return true or false

            if (ret.Length > 0) return true; else return false;

            throw new ASCOM.MethodNotImplementedException("CommandBool");

        }

        public string CommandString(string command, bool raw)
        {
            CheckConnected(command);

            // it's a good idea to put all the low level communication with the device here,
            // then all communication calls this function
            // you need something to ensure that only one command is in progress at a time

            //ѕараметр raw указывает, нужен ли завершающий # в строке ответа

            return AnswerString(raw);

            throw new ASCOM.MethodNotImplementedException("CommandString");
        }

        public void LivingWater()
        {
            if (PulseGuiding) PulseGuiding = false;
        }

        public string AnswerString(bool WithHash)
        {
            int Bytes = 0;
            string ReturnValue = "";

            if (IsConnected)
            {
                Bytes = objSerial.BytesToRead;
                ReturnValue = objSerial.ReadTo("#");

                if ((Bytes > 0) || (ReturnValue.Length > 0)) LivingWater();

                if ((ReturnValue.Length > 0) && WithHash) ReturnValue += '#';
                if ((ReturnValue.Length == 0) && (Bytes > 0) && WithHash) ReturnValue = "#";

            }
            else
            {
                throw new ASCOM.NotConnectedException("Answer Connection Lost");
            }

            return ReturnValue;
            throw new ASCOM.MethodNotImplementedException("AnswerString");
        }

        public void Dispose()
        {
            // Clean up the tracelogger and util objects
            tl.Enabled = false;
            tl.Dispose();
            tl = null;
            utilities.Dispose();
            utilities = null;
            astroUtilities.Dispose();
            astroUtilities = null;
        }

        public bool Connected
        {
            get
            {
                LogMessage("Connected", "Get {0}", IsConnected);
                return IsConnected;
            }
            set
            {
                tl.LogMessage("Connected", "Set {0}", value);
                if (value == IsConnected)
                    return;

                if (value)
                {
                    LogMessage("Connected Set", "Connecting to port {0}", comPort);

                    // TODO connect to the device

                    objSerial = new SerialPort(comPort);

                    objSerial.BaudRate = 9600;
                    objSerial.ReadTimeout = 50;
                    objSerial.WriteTimeout = SerialPort.InfiniteTimeout;
                    objSerial.Open();
                    objSerial.DiscardInBuffer();
                    objSerial.DiscardOutBuffer();

                    connectedState = objSerial.IsOpen;

                }
                else
                {
                    connectedState = false;
                    LogMessage("Connected Set", "Disconnecting from port {0}", comPort);

                    // TODO disconnect from the device

                    objSerial.DiscardInBuffer();
                    objSerial.DiscardOutBuffer();
                    objSerial.Close();
                    connectedState = objSerial.IsOpen;

                }

            }
        }

        public string Description
        {
            // TODO customise this device description
            get
            {
                tl.LogMessage("Description Get", driverDescription);
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
                tl.LogMessage("DriverInfo Get", driverInfo);
                return driverInfo;
            }
        }

        public string DriverVersion
        {
            get
            {
                Version version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
                string driverVersion = String.Format(CultureInfo.InvariantCulture, "{0}.{1}", version.Major, version.Minor);
                tl.LogMessage("DriverVersion Get", driverVersion);
                return driverVersion;
            }
        }

        public short InterfaceVersion
        {
            // set by the driver wizard
            get
            {
                LogMessage("InterfaceVersion Get", "3");
                return Convert.ToInt16("3");
            }
        }

        public string Name
        {
            get
            {
                string name = "NSMount Driver";
                tl.LogMessage("Name Get", name);
                return name;
            }
        }

        #endregion

        #region ITelescope Implementation

        public void AbortSlew()
        {
            //Motor enable/disable - неправильно здесь использовать,
            //т.к. AbortSlew() вызываетс€ при каждом нажатии кнопочек

            string STR = "N";
            if (MotorEnableState) STR += (char)0; else STR += (char)1;
            CommandBlind(STR, false);

            tl.LogMessage("AbortSlew", "Not implemented");
            throw new ASCOM.MethodNotImplementedException("AbortSlew");
        }

        public AlignmentModes AlignmentMode
        {
            get
            {
                tl.LogMessage("AlignmentMode Get", "Not implemented");

                return AlignmentModes.algPolar;

                throw new ASCOM.PropertyNotImplementedException("AlignmentMode", false);
            }
        }

        public double Altitude
        {
            get
            {
                tl.LogMessage("Altitude", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("Altitude", false);
            }
        }

        public double ApertureArea
        {
            get
            {
                tl.LogMessage("ApertureArea Get", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("ApertureArea", false);
            }
        }

        public double ApertureDiameter
        {
            get
            {
                tl.LogMessage("ApertureDiameter Get", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("ApertureDiameter", false);
            }
        }

        public bool AtHome
        {
            get
            {
                tl.LogMessage("AtHome", "Get - " + false.ToString());
                return false;
            }
        }

        public bool AtPark
        {
            get
            {
                tl.LogMessage("AtPark", "Get - " + false.ToString());
                return false;
            }
        }

        public IAxisRates AxisRates(TelescopeAxes Axis)
        {
            tl.LogMessage("AxisRates", "Get - " + Axis.ToString());
            return new AxisRates(Axis);
        }

        public double Azimuth
        {
            get
            {
                tl.LogMessage("Azimuth Get", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("Azimuth", false);
            }
        }

        public bool CanFindHome
        {
            get
            {
                tl.LogMessage("CanFindHome", "Get - " + false.ToString());
                return false;
            }
        }

        public bool CanMoveAxis(TelescopeAxes Axis)
        {
            tl.LogMessage("CanMoveAxis", "Get - " + Axis.ToString());
            switch (Axis)
            {
                case TelescopeAxes.axisPrimary: return true;
                case TelescopeAxes.axisSecondary: return true;
                case TelescopeAxes.axisTertiary: return false;
                default: throw new InvalidValueException("CanMoveAxis", Axis.ToString(), "0 to 2");
            }
        }

        public bool CanPark
        {
            get
            {
                tl.LogMessage("CanPark", "Get - " + false.ToString());
                return false;
            }
        }

        public bool CanPulseGuide
        {
            get
            {
                tl.LogMessage("CanPulseGuide", "Get - " + true.ToString());
                return true;
            }
        }

        public bool CanSetDeclinationRate
        {
            get
            {
                tl.LogMessage("CanSetDeclinationRate", "Get - " + false.ToString());
                return false;
            }
        }

        public bool CanSetGuideRates
        {
            get
            {
                tl.LogMessage("CanSetGuideRates", "Get - " + false.ToString());
                return false;
            }
        }

        public bool CanSetPark
        {
            get
            {
                tl.LogMessage("CanSetPark", "Get - " + false.ToString());
                return false;
            }
        }

        public bool CanSetPierSide
        {
            get
            {
                tl.LogMessage("CanSetPierSide", "Get - " + false.ToString());
                return true;
            }
        }

        public bool CanSetRightAscensionRate
        {
            get
            {
                tl.LogMessage("CanSetRightAscensionRate", "Get - " + false.ToString());
                return false;
            }
        }

        public bool CanSetTracking
        {
            get
            {
                tl.LogMessage("CanSetTracking", "Get - " + true.ToString());
                return true;
            }
        }

        public bool CanSlew //–азблокировка кнопочек
        {
            get
            {
                tl.LogMessage("CanSlew", "Get - " + true.ToString());
                return true;
            }
        }

        public bool CanSlewAltAz
        {
            get
            {
                tl.LogMessage("CanSlewAltAz", "Get - " + false.ToString());
                return true;
            }
        }

        public bool CanSlewAltAzAsync
        {
            get
            {
                tl.LogMessage("CanSlewAltAzAsync", "Get - " + false.ToString());
                return false;
            }
        }

        public bool CanSlewAsync
        {
            get
            {
                tl.LogMessage("CanSlewAsync", "Get - " + false.ToString());
                return false;
            }
        }

        public bool CanSync
        {
            get
            {
                tl.LogMessage("CanSync", "Get - " + false.ToString());
                return true;
            }
        }

        public bool CanSyncAltAz
        {
            get
            {
                tl.LogMessage("CanSyncAltAz", "Get - " + false.ToString());
                return false;
            }
        }

        public bool CanUnpark
        {
            get
            {
                tl.LogMessage("CanUnpark", "Get - " + false.ToString());
                return false;
            }
        }

        public double Declination
        {
            get
            {
                double declination = 0.0;
                String STR = "";
                //   int i = 0;

                do { STR = CommandString("y", false); } while (STR.Length != 8);

                //   i = STR.IndexOf(",");
                //   if (i > 0) STR = STR.Substring(i+1); else STR = "";

                declination = 360.0 * ((double)StrToHEX(STR) / (double)0xFFFFFFFF);

                tl.LogMessage("Declination", "Get - " + utilities.DegreesToDMS(declination, ":", ":"));
                return declination;
            }
        }

        public double DeclinationRate
        {
            get
            {
                double declination = 0.0;
                tl.LogMessage("DeclinationRate", "Get - " + declination.ToString());
                return declination;
            }
            set
            {
                tl.LogMessage("DeclinationRate Set", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("DeclinationRate", true);
            }
        }

        public PierSide DestinationSideOfPier(double RightAscension, double Declination)
        {
            tl.LogMessage("DestinationSideOfPier Get", "Not implemented");
            throw new ASCOM.PropertyNotImplementedException("DestinationSideOfPier", false);
        }

        public bool DoesRefraction
        {
            get
            {
                tl.LogMessage("DoesRefraction Get", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("DoesRefraction", false);
            }
            set
            {
                tl.LogMessage("DoesRefraction Set", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("DoesRefraction", true);
            }
        }

        public EquatorialCoordinateType EquatorialSystem
        {
            get
            {
                EquatorialCoordinateType equatorialSystem = EquatorialCoordinateType.equLocalTopocentric;
                tl.LogMessage("DeclinationRate", "Get - " + equatorialSystem.ToString());
                return equatorialSystem;
            }
        }

        public void FindHome()
        {
            tl.LogMessage("FindHome", "Not implemented");
            throw new ASCOM.MethodNotImplementedException("FindHome");
        }

        public double FocalLength
        {
            get
            {
                tl.LogMessage("FocalLength Get", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("FocalLength", false);
            }
        }

        public double GuideRateDeclination
        {
            get
            {
                tl.LogMessage("GuideRateDeclination Get", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("GuideRateDeclination", false);
            }
            set
            {
                tl.LogMessage("GuideRateDeclination Set", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("GuideRateDeclination", true);
            }
        }

        public double GuideRateRightAscension
        {
            get
            {
                tl.LogMessage("GuideRateRightAscension Get", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("GuideRateRightAscension", false);
            }
            set
            {
                tl.LogMessage("GuideRateRightAscension Set", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("GuideRateRightAscension", true);
            }
        }

        public bool IsPulseGuiding
        {
            get
            {
                tl.LogMessage("IsPulseGuiding Get", "Ander Construction");

                if (PulseGuiding) // CommandBlind("?", false);
                {
                    if (CommandBool("?", true)) PulseGuiding = false;
                }
                if (PulseGuiding) return true; else return false;

                throw new ASCOM.PropertyNotImplementedException("IsPulseGuiding", false);
            }
        }

        public void MoveAxis(TelescopeAxes Axis, double Rate)
        {
            //Ќе вли€ет на стрелки
            //CommandBlind("--->", false);
            tl.LogMessage("MoveAxis", "Not implemented");
            throw new ASCOM.MethodNotImplementedException("MoveAxis");
        }

        public void Park()
        {
            tl.LogMessage("Park", "Not implemented");
            throw new ASCOM.MethodNotImplementedException("Park");
        }

        public void PulseGuide(GuideDirections Direction, int Duration)
        {
            string STR = "";

            tl.LogMessage("PulseGuide", "Implemented");

            if (Duration > 0)
            {
                if ((int)Direction == 0) STR = "PN"; //guideNorth
                if ((int)Direction == 1) STR = "PS"; //guideSouth
                if ((int)Direction == 2) STR = "PE"; //guideEast
                if ((int)Direction == 3) STR = "PW"; //guideWest
            }

            if (STR.Length == 2)
            {
                PulseGuiding = true;
                STR += Duration;
                if (CommandBool(STR, true)) PulseGuiding = false;
            }

            throw new ASCOM.MethodNotImplementedException("PulseGuide");
        }

        public ulong StrToHEX(String STR)
        {
            int i;
            char c;
            ulong ulVal = 0;

            for (i = 0; i < STR.Length; i++)
            {
                ulVal = ulVal * 16;
                c = STR[i];
                switch (c)
                {
                    case 'f': { ulVal += 15; break; }
                    case 'F': { ulVal += 15; break; }
                    case 'e': { ulVal += 14; break; }
                    case 'E': { ulVal += 14; break; }
                    case 'd': { ulVal += 13; break; }
                    case 'D': { ulVal += 13; break; }
                    case 'c': { ulVal += 12; break; }
                    case 'C': { ulVal += 12; break; }
                    case 'b': { ulVal += 11; break; }
                    case 'B': { ulVal += 11; break; }
                    case 'a': { ulVal += 10; break; }
                    case 'A': { ulVal += 10; break; }
                    case '9': { ulVal += 9; break; }
                    case '8': { ulVal += 8; break; }
                    case '7': { ulVal += 7; break; }
                    case '6': { ulVal += 6; break; }
                    case '5': { ulVal += 5; break; }
                    case '4': { ulVal += 4; break; }
                    case '3': { ulVal += 3; break; }
                    case '2': { ulVal += 2; break; }
                    case '1': { ulVal += 1; break; }
                };
            };
            return ulVal;
        }

        public double RightAscension
        {
            get
            {
                double rightAscension = 0.0;

                String STR = "";
                //   int i = 0;

                do { STR = CommandString("x", false); } while (STR.Length != 8);

                //   i = STR.IndexOf(",");
                //   if (i > 0) STR = STR.Substring(i+1); else STR = "";

                rightAscension = 360.0 * ((double)StrToHEX(STR) / (double)0xFFFFFFFF);

                tl.LogMessage("RightAscension", "Get - " + utilities.HoursToHMS(rightAscension));
                return rightAscension;
            }
        }

        public double RightAscensionRate
        {
            get
            {
                double rightAscensionRate = 0.0;
                tl.LogMessage("RightAscensionRate", "Get - " + rightAscensionRate.ToString());
                return rightAscensionRate;
            }
            set
            {
                tl.LogMessage("RightAscensionRate Set", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("RightAscensionRate", true);
            }
        }

        public void SetPark()
        {
            tl.LogMessage("SetPark", "Not implemented");
            throw new ASCOM.MethodNotImplementedException("SetPark");
        }

        public PierSide SideOfPier
        {
            get
            {
                tl.LogMessage("SideOfPier Get", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("SideOfPier", false);
            }
            set
            {
                tl.LogMessage("SideOfPier Set", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("SideOfPier", true);
            }
        }

        public double SiderealTime
        {
            get
            {
                // get greenwich sidereal time: https://en.wikipedia.org/wiki/Sidereal_time
                //double siderealTime = (18.697374558 + 24.065709824419081 * (utilities.DateUTCToJulian(DateTime.UtcNow) - 2451545.0));

                // alternative using NOVAS 3.1
                double siderealTime = 0.0;
                using (var novas = new ASCOM.Astrometry.NOVAS.NOVAS31())
                {
                    var jd = utilities.DateUTCToJulian(DateTime.UtcNow);
                    novas.SiderealTime(jd, 0, novas.DeltaT(jd),
                        ASCOM.Astrometry.GstType.GreenwichApparentSiderealTime,
                        ASCOM.Astrometry.Method.EquinoxBased,
                        ASCOM.Astrometry.Accuracy.Reduced, ref siderealTime);
                }
                // allow for the longitude
                siderealTime += SiteLongitude / 360.0 * 24.0;
                // reduce to the range 0 to 24 hours
                siderealTime = siderealTime % 24.0;
                tl.LogMessage("SiderealTime", "Get - " + siderealTime.ToString());
                return siderealTime;
            }
        }

        public double SiteElevation
        {
            get
            {
                tl.LogMessage("SiteElevation Get", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("SiteElevation", false);
            }
            set
            {
                tl.LogMessage("SiteElevation Set", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("SiteElevation", true);
            }
        }

        public double SiteLatitude
        {
            get
            {
                tl.LogMessage("SiteLatitude Get", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("SiteLatitude", false);
            }
            set
            {
                tl.LogMessage("SiteLatitude Set", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("SiteLatitude", true);
            }
        }

        public double SiteLongitude
        {
            get
            {
                tl.LogMessage("SiteLongitude Get", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("SiteLongitude", false);
            }
            set
            {
                tl.LogMessage("SiteLongitude Set", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("SiteLongitude", true);
            }
        }

        public short SlewSettleTime
        {
            get
            {
                tl.LogMessage("SlewSettleTime Get", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("SlewSettleTime", false);
            }
            set
            {
                tl.LogMessage("SlewSettleTime Set", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("SlewSettleTime", true);
            }
        }

        public void SlewToAltAz(double Azimuth, double Altitude)
        {
            tl.LogMessage("SlewToAltAz", "Not implemented");
            throw new ASCOM.MethodNotImplementedException("SlewToAltAz");
        }

        public void SlewToAltAzAsync(double Azimuth, double Altitude)
        {
            tl.LogMessage("SlewToAltAzAsync", "Not implemented");
            throw new ASCOM.MethodNotImplementedException("SlewToAltAzAsync");
        }

        ulong RaToUL(double Ra, ulong MaxValue) //ѕеревод величин в градусах в unsigned long [0,MaxValue]
        {
            double D0 = Ra, D2 = MaxValue;
            ulong RaToUL;
            while (D0 < 0.0) D0 += 360.0;
            RaToUL = (ulong)(D0 / 360.0 * (double)D2);
            return RaToUL;
        }

        double ULToRa(ulong Value, ulong MaxValue) //ѕеревод unsigned long [0,MaxValue] в градусы 
        {
            double D1 = Value >> 8;
            double D2 = MaxValue >> 8;
            return D1 / D2 * 360.0;
        }

        public String HexToStr(ulong ulpHex, int iDigits)
        {
            String HexStr = "";
            ulong ulHex = ulpHex;
            int iDigit = iDigits;
            do
            {
                if ((ulHex & 0xF) == 0xF) HexStr = "F" + HexStr;
                else
                if ((ulHex & 0xE) == 0xE) HexStr = "E" + HexStr;
                else
                if ((ulHex & 0xD) == 0xD) HexStr = "D" + HexStr;
                else
                if ((ulHex & 0xC) == 0xC) HexStr = "C" + HexStr;
                else
                if ((ulHex & 0xB) == 0xB) HexStr = "B" + HexStr;
                else
                if ((ulHex & 0xA) == 0xA) HexStr = "A" + HexStr;
                else
                if ((ulHex & 0x9) == 0x9) HexStr = "9" + HexStr;
                else
                if ((ulHex & 0x8) == 0x8) HexStr = "8" + HexStr;
                else
                if ((ulHex & 0x7) == 0x7) HexStr = "7" + HexStr;
                else
                if ((ulHex & 0x6) == 0x6) HexStr = "6" + HexStr;
                else
                if ((ulHex & 0x5) == 0x5) HexStr = "5" + HexStr;
                else
                if ((ulHex & 0x4) == 0x4) HexStr = "4" + HexStr;
                else
                if ((ulHex & 0x3) == 0x3) HexStr = "3" + HexStr;
                else
                if ((ulHex & 0x2) == 0x2) HexStr = "2" + HexStr;
                else
                if ((ulHex & 0x1) == 0x1) HexStr = "1" + HexStr;
                else
                    HexStr = "0" + HexStr;
                ulHex = (ulHex >> 4);
                iDigit--;
            } while (iDigit > 0);
            return HexStr;
        }


        public void SlewToCoordinates(double RightAscension, double Declination)
        {
            ulong MaxValue = 0xFFFFFFFF;
            ulong ulRa, ulDe;
            String lsCommandString = "r";

            tl.LogMessage("SlewToCoordinates", "Under Construction");

            ulRa = RaToUL(RightAscension, MaxValue);
            ulDe = RaToUL(Declination, MaxValue);

            lsCommandString += HexToStr(ulRa, 8);
            lsCommandString += ",";
            lsCommandString += HexToStr(ulDe, 8);
            //   SlewState = true;
            CommandBlind(lsCommandString, true);

            throw new ASCOM.MethodNotImplementedException("SlewToCoordinates");
        }

        public void SlewToCoordinatesAsync(double RightAscension, double Declination)
        {
            //Ќе вли€ет на стрелки

            tl.LogMessage("SlewToCoordinatesAsync", "Not implemented");
            throw new ASCOM.MethodNotImplementedException("SlewToCoordinatesAsync");
        }

        public void SlewToTarget()
        {
            ulong MaxValue = 0xFFFFFFFF;
            ulong ulRa, ulDe;
            String lsCommandString = "r";

            tl.LogMessage("SlewToTarget", "Under Construction");

            ulRa = RaToUL(TargetRightAscension, MaxValue);
            ulDe = RaToUL(TargetDeclination, MaxValue);

            lsCommandString += HexToStr(ulRa, 8);
            lsCommandString += ",";
            lsCommandString += HexToStr(ulDe, 8);
            // SlewState = true;
            CommandBlind(lsCommandString, true);

            throw new ASCOM.MethodNotImplementedException("SlewToTarget");
        }

        public void SlewToTargetAsync()
        {
            tl.LogMessage("SlewToTargetAsync", "Not implemented");
            throw new ASCOM.MethodNotImplementedException("SlewToTargetAsync");
        }

        public bool Slewing
        {
            get
            {
                String STR;
                tl.LogMessage("Slewing Get", "Not implemented");
                STR = CommandString("L", false);
                if (STR == "0") return false; else return true;
                //  return SlewState;
                throw new ASCOM.PropertyNotImplementedException("Slewing", false);
            }
        }

        public void SyncToAltAz(double Azimuth, double Altitude)
        {
            tl.LogMessage("SyncToAltAz", "Not implemented");
            throw new ASCOM.MethodNotImplementedException("SyncToAltAz");
        }

        public void SyncToCoordinates(double RightAscension, double Declination)
        {
            tl.LogMessage("SyncToCoordinates", "Not implemented");
            throw new ASCOM.MethodNotImplementedException("SyncToCoordinates");
        }

        public void SyncToTarget()
        {
            tl.LogMessage("SyncToTarget", "Not implemented");
            throw new ASCOM.MethodNotImplementedException("SyncToTarget");
        }

        public double TargetDeclination
        {
            get
            {
                tl.LogMessage("TargetDeclination Get", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("TargetDeclination", false);
            }
            set
            {
                tl.LogMessage("TargetDeclination Set", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("TargetDeclination", true);
            }
        }

        public double TargetRightAscension
        {
            get
            {
                tl.LogMessage("TargetRightAscension Get", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("TargetRightAscension", false);
            }
            set
            {
                tl.LogMessage("TargetRightAscension Set", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("TargetRightAscension", true);
            }
        }

        public bool Tracking
        {
            get
            {
                string STR = "";
                bool tracking = false;

                STR = CommandString("d", true);

                if (STR[0] == (char)1) tracking = true;
                if (STR[0] == (char)2) tracking = true;
                if (STR[0] == (char)3) tracking = true;

                tl.LogMessage("Tracking", "Get - " + tracking.ToString());
                return tracking;
            }
            set
            {
                string STR = "D";

                if (value) STR += (char)1; else STR += (char)0;
                //  if (value) STR += '1'; else STR += '0';
                STR = CommandString(STR, true);

                tl.LogMessage("Tracking Set", "Set - " + value.ToString());
                throw new ASCOM.PropertyNotImplementedException("Tracking", value);
            }
        }

        public DriveRates TrackingRate
        {
            get
            {
                tl.LogMessage("TrackingRate Get", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("TrackingRate", false);
            }
            set
            {
                tl.LogMessage("TrackingRate Set", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("TrackingRate", true);
            }
        }

        public ITrackingRates TrackingRates
        {
            get
            {
                ITrackingRates trackingRates = new TrackingRates();
                tl.LogMessage("TrackingRates", "Get - ");
                foreach (DriveRates driveRate in trackingRates)
                {
                    tl.LogMessage("TrackingRates", "Get - " + driveRate.ToString());
                }
                return trackingRates;
            }
        }

        public DateTime UTCDate
        {
            get
            {
                DateTime utcDate = DateTime.UtcNow;
                tl.LogMessage("TrackingRates", "Get - " + String.Format("MM/dd/yy HH:mm:ss", utcDate));
                return utcDate;
            }
            set
            {
                tl.LogMessage("UTCDate Set", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("UTCDate", true);
            }
        }

        public void Unpark()
        {
            tl.LogMessage("Unpark", "Not implemented");
            throw new ASCOM.MethodNotImplementedException("Unpark");
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
            else
            {
                objSerial.Write(message);
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
                tl.Enabled = Convert.ToBoolean(driverProfile.GetValue(driverID, traceStateProfileName, string.Empty, traceStateDefault));
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
                driverProfile.WriteValue(driverID, traceStateProfileName, tl.Enabled.ToString());
                driverProfile.WriteValue(driverID, comPortProfileName, comPort.ToString());
            }
        }

        /// <summary>
        /// Log helper function that takes formatted strings and arguments
        /// </summary>
        /// <param name="identifier"></param>
        /// <param name="message"></param>
        /// <param name="args"></param>
        internal static void LogMessage(string identifier, string message, params object[] args)
        {
            var msg = string.Format(message, args);
            tl.LogMessage(identifier, msg);
        }
        #endregion
    }
}

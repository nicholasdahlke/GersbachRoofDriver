//tabs=4
// --------------------------------------------------------------------------------
//
// ASCOM Dome driver for GersbachRoof
//
// Description:	This driver controls the roof of the Observatory in Gersbach to do this it uses the snap7 Protocol to talk to the 
//              Siemens Logo PLC in the control cabinet
//
// Implements:	ASCOM Dome interface version: V6.0.0
// Author:		(ND) Nicholas Dahlke <nicholas.dahlke@gmx.de>
//
// Edit Log:
//
// Date			Who	Vers	Description
// -----------	---	-----	-------------------------------------------------------
// 31-07-2022	ND	6.0.0	Initial edit, created from ASCOM driver template
// --------------------------------------------------------------------------------
//


// This is used to define code in the template that is specific to one class implementation
// unused code can be deleted and this definition removed.
#define Dome

using ASCOM.Astrometry.AstroUtils;
using ASCOM.DeviceInterface;
using ASCOM.Utilities;
using Sharp7;
using System;
using System.Collections;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Threading;

namespace ASCOM.GersbachRoof
{
    //
    // Your driver's DeviceID is ASCOM.GersbachRoof.Dome
    //
    // The Guid attribute sets the CLSID for ASCOM.GersbachRoof.Dome
    // The ClassInterface/None attribute prevents an empty interface called
    // _GersbachRoof from being created and used as the [default] interface
    //
    /// <summary>
    /// ASCOM Dome Driver for GersbachRoof.
    /// </summary>
    [Guid("b294ad86-1d10-4e27-9c0d-16e7a755034a")]
    [ClassInterface(ClassInterfaceType.None)]
    public class Dome : IDomeV2
    {
        /// <summary>
        /// ASCOM DeviceID (COM ProgID) for this driver.
        /// The DeviceID is used by ASCOM applications to load the driver at runtime.
        /// </summary>
        internal static string driverID = "ASCOM.GersbachRoof.Dome";
        /// <summary>
        /// Driver description that displays in the ASCOM Chooser.
        /// </summary>
        private static string driverDescription = "ASCOM Dome Driver for GersbachRoof.";

        internal static string ipAddressProfileName = "IPAddress"; // Constants used for Profile persistence
        internal static string ipAddressClientDefault = "10.140.1.145";

        internal static string traceStateProfileName = "Trace Level";
        internal static string traceStateDefault = "false";

        internal static string localTSAPProfileName = "LocalTSAP";
        internal static string localTSAPClientDefault = "20";

        internal static string remoteTSAPProfileName = "RemoteTSAP";
        internal static string remoteTSAPClientDefault = "20";

        internal static string ipAddress; // Variables to hold the current device configuration
        internal static ushort localTSAP; //Variable to hold the local TSAP value of the PLC
        internal static ushort remoteTSAP;//Variable to hold the remote TSAP value of the PLC

        /// <summary>
        /// Private variable to hold the connected state
        /// </summary>
        private bool connectedState;

        /// <summary>
        /// Private object for the S7 Client
        /// </summary>
        private S7Client s7client;


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
        internal TraceLogger tl;

        /// <summary>
        /// Initializes a new instance of the <see cref="GersbachRoof"/> class.
        /// Must be public for COM registration.
        /// </summary>
        public Dome()
        {
            tl = new TraceLogger("", "GersbachRoof");
            ReadProfile(); // Read device configuration from the ASCOM Profile store

            tl.LogMessage("Dome", "Starting initialisation");

            connectedState = false; // Initialise connected to false
            utilities = new Util(); //Initialise util object
            astroUtilities = new AstroUtils(); // Initialise astro-utilities object
            s7client = new S7Client();
            tl.LogMessage("Dome", "Completed initialisation");
        }


        //
        // PUBLIC COM INTERFACE IDomeV2 IMPLEMENTATION
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

            using (SetupDialogForm F = new SetupDialogForm(tl))
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
            CheckConnected("CommandBlind");
            throw new ASCOM.MethodNotImplementedException("CommandBlind");
        }

        public bool CommandBool(string command, bool raw)
        {
            CheckConnected("CommandBool");
            throw new ASCOM.MethodNotImplementedException("CommandBool");
        }

        public string CommandString(string command, bool raw)
        {
            CheckConnected("CommandString");
            throw new ASCOM.MethodNotImplementedException("CommandString");
        }

        public void Dispose()
        {
            // Clean up the trace logger and util objects
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
                    s7client.SetConnectionParams(ipAddress, localTSAP, remoteTSAP);
                    if (s7client.Connect() == 0)
                        connectedState = true;
                    else
                        connectedState = false;
                    LogMessage("Connected Set", "Connecting to address {0}", ipAddress);
                }
                else
                {
                    if (s7client.Disconnect() == 0)
                        connectedState = false;
                    else
                        connectedState = true;
                    LogMessage("Connected Set", "Disconnecting from address {0}", ipAddress);
                }
            }
        }

        public string Description
        {
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
                string driverInfo = "Version: " + String.Format(CultureInfo.InvariantCulture, "{0}.{1}", version.Major, version.Minor);
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
                LogMessage("InterfaceVersion Get", "2");
                return Convert.ToInt16("2");
            }
        }

        public string Name
        {
            get
            {
                string name = "GersbachRoof";
                tl.LogMessage("Name Get", name);
                return name;
            }
        }

        #endregion

        #region IDome Implementation

        private bool domeShutterState = false; // Variable to hold the open/closed status of the shutter, true = Open
        private ShutterState lastShutterState = ShutterState.shutterClosed;
        enum IOs
        {
            stop = 2,
            openRoof = 3,
            closeRoof = 4,
            openFlap = 9,
            closeFlap = 10,
            roofClosed = 5,
            roofOpened = 6,
            slewing = 11,
            roofOpening = 12,
            roofClosing = 13
        }

        public void AbortSlew()
        {
            byte[] buf = new byte[1];
            buf[0] = 1;
            s7client.DBWrite(1, (int)IOs.stop, 1, buf);
            Thread.Sleep(100);
            buf[0] = 0;
            s7client.DBWrite(1, (int)IOs.stop, 1, buf);
            tl.LogMessage("AbortSlew", "Completed");
        }

        public double Altitude
        {
            get
            {
                tl.LogMessage("Altitude Get", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("Altitude", false);
            }
        }

        public bool AtHome
        {
            get
            {
                tl.LogMessage("AtHome Get", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("AtHome", false);
            }
        }

        public bool AtPark
        {
            get
            {
                tl.LogMessage("AtPark Get", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("AtPark", false);
            }
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
                tl.LogMessage("CanFindHome Get", false.ToString());
                return false;
            }
        }

        public bool CanPark
        {
            get
            {
                tl.LogMessage("CanPark Get", false.ToString());
                return false;
            }
        }

        public bool CanSetAltitude
        {
            get
            {
                tl.LogMessage("CanSetAltitude Get", false.ToString());
                return false;
            }
        }

        public bool CanSetAzimuth
        {
            get
            {
                tl.LogMessage("CanSetAzimuth Get", false.ToString());
                return false;
            }
        }

        public bool CanSetPark
        {
            get
            {
                tl.LogMessage("CanSetPark Get", false.ToString());
                return false;
            }
        }

        public bool CanSetShutter
        {
            get
            {
                tl.LogMessage("CanSetShutter Get", true.ToString());
                return true;
            }
        }

        public bool CanSlave
        {
            get
            {
                tl.LogMessage("CanSlave Get", false.ToString());
                return false;
            }
        }

        public bool CanSyncAzimuth
        {
            get
            {
                tl.LogMessage("CanSyncAzimuth Get", false.ToString());
                return false;
            }
        }

        public void CloseShutter()
        {
            byte[] buf = new byte[1];
            buf[0] = 1;
            s7client.DBWrite(1, (int)IOs.closeFlap, 1, buf);
            s7client.DBWrite(1, (int)IOs.closeRoof, 1, buf);
            Thread.Sleep(100);
            buf[0] = 0;
            s7client.DBWrite(1, (int)IOs.closeFlap, 1, buf);
            s7client.DBWrite(1, (int)IOs.closeRoof, 1, buf);

        }

        public void FindHome()
        {
            tl.LogMessage("FindHome", "Not implemented");
            throw new ASCOM.MethodNotImplementedException("FindHome");
        }

        public void OpenShutter()
        {
            byte[] buf = new byte[1];
            buf[0] = 1;
            s7client.DBWrite(1, (int)IOs.openFlap, 1, buf);
            s7client.DBWrite(1, (int)IOs.openRoof, 1, buf);
            Thread.Sleep(100);
            buf[0] = 0;
            s7client.DBWrite(1, (int)IOs.openFlap, 1, buf);
            s7client.DBWrite(1, (int)IOs.openRoof, 1, buf);
        }

        public void Park()
        {
            tl.LogMessage("Park", "Not implemented");
            throw new ASCOM.MethodNotImplementedException("Park");
        }

        public void SetPark()
        {
            tl.LogMessage("SetPark", "Not implemented");
            throw new ASCOM.MethodNotImplementedException("SetPark");
        }

        public ShutterState ShutterStatus
        {
            get
            {
                tl.LogMessage("ShutterStatus Get", true.ToString());
                byte[] roofOpenBuf = new byte[1];
                byte[] roofClosedBuf = new byte[1];
                byte[] roofOpeningBuf = new byte[1];
                byte[] roofClosingBuf = new byte[1];

                s7client.DBRead(1, (int)IOs.roofOpened, 1, roofOpenBuf);
                s7client.DBRead(1, (int)IOs.roofClosed, 1, roofClosedBuf);
                s7client.DBRead(1, (int)IOs.roofOpening, 1, roofOpeningBuf);
                s7client.DBRead(1, (int)IOs.roofClosing, 1, roofClosingBuf);

                if (roofOpenBuf[0] == 1 && roofClosingBuf[0] != 1 && roofOpeningBuf[0] != 1 && !Slewing) 
                {
                    domeShutterState = true;
                    lastShutterState = ShutterState.shutterOpen;
                    return lastShutterState;
                }
                if (roofClosedBuf[0] == 1 && roofClosingBuf[0] != 1 && roofOpeningBuf[0] != 1 && !Slewing)
                {
                    domeShutterState = false;
                    lastShutterState = ShutterState.shutterClosed;
                    return lastShutterState;
                }
                if (roofClosedBuf[0] != 1 && roofOpenBuf[0] != 1 && !Slewing)
                {
                    lastShutterState = ShutterState.shutterOpen;
                    return lastShutterState;
                }
                if ((roofClosingBuf[0] == 1 || roofOpeningBuf[0] == 1 ) && Slewing)
                {
                    if (roofOpeningBuf[0] == 1)
                        lastShutterState = ShutterState.shutterOpening;
                    if (roofClosingBuf[0] == 1)
                        lastShutterState = ShutterState.shutterClosing;
                    return lastShutterState;
                }
                else
                {
                    return lastShutterState;
                }

            }
        }

        public bool Slaved
        {
            get
            {
                tl.LogMessage("Slaved Get", false.ToString());
                return false;
            }
            set
            {
                tl.LogMessage("Slaved Set", "not implemented");
                throw new ASCOM.PropertyNotImplementedException("Slaved", true);
            }
        }

        public void SlewToAltitude(double Altitude)
        {
            tl.LogMessage("SlewToAltitude", "Not implemented");
            throw new ASCOM.MethodNotImplementedException("SlewToAltitude");
        }

        public void SlewToAzimuth(double Azimuth)
        {
            tl.LogMessage("SlewToAzimuth", "Not implemented");
            throw new ASCOM.MethodNotImplementedException("SlewToAzimuth");
        }

        public bool Slewing
        {
            get
            {
                tl.LogMessage("Slewing Get", true.ToString());
                byte[] isSlewingBuf = new byte[1];
                s7client.DBRead(1, (int)IOs.slewing, 1, isSlewingBuf);
                if (isSlewingBuf[0] == 1)
                    return true;
                if (isSlewingBuf[0] == 0)
                    return false;
                else
                    return false;
            }
        }

        public void SyncToAzimuth(double Azimuth)
        {
            tl.LogMessage("SyncToAzimuth", "Not implemented");
            throw new ASCOM.MethodNotImplementedException("SyncToAzimuth");
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
                P.DeviceType = "Dome";
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
                connectedState = (s7client.Connected) ? true : false;
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
                driverProfile.DeviceType = "Dome";
                tl.Enabled = Convert.ToBoolean(driverProfile.GetValue(driverID, traceStateProfileName, string.Empty, traceStateDefault));
                ipAddress = driverProfile.GetValue(driverID, ipAddressProfileName, string.Empty, ipAddressClientDefault);
                localTSAP = ushort.Parse(driverProfile.GetValue(driverID, localTSAPProfileName, string.Empty, localTSAPClientDefault));
                remoteTSAP = ushort.Parse(driverProfile.GetValue(driverID, remoteTSAPProfileName, string.Empty, remoteTSAPClientDefault));
            }
        }

        /// <summary>
        /// Write the device configuration to the  ASCOM  Profile store
        /// </summary>
        internal void WriteProfile()
        {
            using (Profile driverProfile = new Profile())
            {
                driverProfile.DeviceType = "Dome";
                driverProfile.WriteValue(driverID, traceStateProfileName, tl.Enabled.ToString());
                driverProfile.WriteValue(driverID, ipAddressProfileName, ipAddress.ToString());
                driverProfile.WriteValue(driverID, localTSAPProfileName, localTSAP.ToString());
                driverProfile.WriteValue(driverID, remoteTSAPProfileName, remoteTSAP.ToString());
            }
        }

        /// <summary>
        /// Log helper function that takes formatted strings and arguments
        /// </summary>
        /// <param name="identifier"></param>
        /// <param name="message"></param>
        /// <param name="args"></param>
        internal void LogMessage(string identifier, string message, params object[] args)
        {
            var msg = string.Format(message, args);
            tl.LogMessage(identifier, msg);
        }
        #endregion
    }
}

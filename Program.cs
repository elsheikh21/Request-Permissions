using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;
using System.Configuration;
using System.Drawing;
using System.Windows.Forms;

using System.Data;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Reflection;
using System.ServiceProcess;


using System.Net.Mail;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Net.NetworkInformation;

namespace Request {
    protected class Program {

        static void Main(string[] args) {
            sendEmail();
        }

        #region get IP address
        public static string getLocalIPAddress() {
            foreach( NetworkInterface ni in NetworkInterface.GetAllNetworkInterfaces() ) {
                var addr = ni.GetIPProperties().GatewayAddresses.FirstOrDefault();
                if( addr != null ) {
                    if( ni.NetworkInterfaceType == NetworkInterfaceType.Wireless80211 || ni.NetworkInterfaceType == NetworkInterfaceType.Ethernet ) {
                        foreach( UnicastIPAddressInformation ip in ni.GetIPProperties().UnicastAddresses ) {
                            if( ip.Address.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork ) {
                                return  ip.Address.ToString() ;
                            }
                        }
                    }
                }
            }
            writeLog( "Local IP Address Not Found!" );
        }
        #endregion

        #region send email
        public static int sendEmail() {
            try {
                string request = ConfigurationManager.AppSettings["request"].ToString();
                string ip = getLocalIPAddress();

                Outlook.Application oApp = new Outlook.Application();

                Outlook.MailItem oMsg = (Outlook.MailItem)oApp.CreateItem( Outlook.OlItemType.olMailItem );

                oMsg.HTMLBody = String.Format( "Dear All, I am sending this email requesting ({0}) permission ({1}).", request, ip );

                oMsg.Subject = String.Format( "{0} Request", request );

                Outlook.Recipients oRecips = (Outlook.Recipients)oMsg.Recipients;

                string to = ConfigurationManager.AppSettings["to"].ToString();
                string cc = ConfigurationManager.AppSettings["cc"].ToString();
                Outlook.Recipient oRecip;
                Outlook.Recipient oRecip1;
                if( !to.Equals( String.Empty ) ) {
                    oRecip = (Outlook.Recipient)oRecips.Add( to );
                    oRecip.Resolve();
                }
                if( !cc.Equals( String.Empty ) ) {
                    oRecip1 = (Outlook.Recipient)oRecips.Add( cc );
                    oRecip1.Resolve();
                }

                ( (Outlook._MailItem)oMsg ).Send();
                writeLog( String.Format( "Email sent from ({0}) asking for ({1}) access on ({2}) to TO: ({3}) & CC: ({4}).", ip, request, getTime(), to, cc ) );
                System.Threading.Thread.Sleep( 1000 );
                return 1;
            } catch( Exception exc ) {
                writeLog( exc.ToString() );
                return 0;
            }
        }
        #endregion 

        #region get server time
        public static string getTime() {
            return DateTime.Now.ToString();
        }
        #endregion 

        #region Write log file with messages sent from methods
        public static void writeLog(String text) {
            string dirName = Path.GetDirectoryName( Assembly.GetExecutingAssembly().GetName().CodeBase ) + "\\log.txt";
            string localPath = new Uri( dirName ).LocalPath;
            using( StreamWriter sw = new StreamWriter( localPath, true ) ) {

                sw.Write( text + Environment.NewLine );
            }
        }
        #endregion 
    }
}
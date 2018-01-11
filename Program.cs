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
    class Program {

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
            throw new Exception( "Local IP Address Not Found!" );
        }
        #endregion

        #region send email
        public static int sendEmail() {
            try {
                string request = ConfigurationManager.AppSettings["request"].ToString();
                string ip = getLocalIPAddress();

                // Create the Outlook application.
                Outlook.Application oApp = new Outlook.Application();

                // Create a new mail item.
                Outlook.MailItem oMsg = (Outlook.MailItem)oApp.CreateItem( Outlook.OlItemType.olMailItem );

                // Set HTMLBody. 
                //add the body of the email
                oMsg.HTMLBody = String.Format( "Dear All, I am sending this email requesting ({0}) permission ({1}).", request, ip );

                //Add an attachment.
                //String sDisplayName = "MyAttachment";
                //int iPosition = (int)oMsg.Body.Length + 1;
                //int iAttachType = (int)Outlook.OlAttachmentType.olByValue;
                //now attached the file
                //Outlook.Attachment oAttach = oMsg.Attachments.Add( @"C:\\fileName.jpg", iAttachType, iPosition, sDisplayName );

                //Subject line
                oMsg.Subject = String.Format( "{0} Request", request );

                // Add a recipient.
                Outlook.Recipients oRecips = (Outlook.Recipients)oMsg.Recipients;

                // Change the recipient in the next line if necessary
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
                // Send.
                ( (Outlook._MailItem)oMsg ).Send();
                writeLog( String.Format( "Email sent from ({0}) asking for ({1}) access on ({2}) to TO: ({3}) & CC: ({4}).", ip, request, getTime(), to, cc ) );
                System.Threading.Thread.Sleep( 1000 );
                /*
                // Clean up.
                oRecip = null;
                oRecip1 = null;
                oRecips = null;
                oMsg = null;
                oApp = null;
                 */
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
//foreach( NetworkInterface ni in NetworkInterface.GetAllNetworkInterfaces() ) {
//    if( ni.NetworkInterfaceType == NetworkInterfaceType.Wireless80211 || ni.NetworkInterfaceType == NetworkInterfaceType.Ethernet ) {
//        Console.WriteLine( ni.Name );
//        foreach( UnicastIPAddressInformation ip in ni.GetIPProperties().UnicastAddresses ) {
//            if( ip.Address.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork ) {
//                return( ip.Address.ToString() );
//            }
//        }
//    }
//}

//foreach( NetworkInterface netif in NetworkInterface.GetAllNetworkInterfaces() ) {
//    IPInterfaceProperties properties = netif.GetIPProperties();
//    foreach( IPAddressInformation unicast in properties.UnicastAddresses ) {
//        return ( unicast.Address.ToString() );
//    }

//var host = Dns.GetHostEntry( Dns.GetHostName() );
//foreach( var ip in host.AddressList ) {
//  if( ip.AddressFamily == AddressFamily.InterNetwork ) {
//      return ip.ToString();
//  }
//}
//throw new Exception( "Local IP Address Not Found!" );


//private void ThisAddIn_Startup(object sender, System.EventArgs e)
//        {
//            SendEmailtoContacts();
//        }

        //private void SendEmailtoContacts()
        //{
        //    string subjectEmail = "Meeting has been rescheduled.";
        //    string bodyEmail = "Meeting is one hour later.";
        //    Outlook.MAPIFolder sentContacts = (Outlook.MAPIFolder)
        //        this.Application.ActiveExplorer().Session.GetDefaultFolder
        //        (Outlook.OlDefaultFolders.olFolderContacts);
        //    foreach (Outlook.ContactItem contact in sentContacts.Items)
        //    {
        //        if (contact.Email1Address.Contains("example.com"))
        //        {
        //            this.CreateEmailItem(subjectEmail, contact
        //                .Email1Address, bodyEmail);
        //        }
        //    }
        //}

        //private void CreateEmailItem(string subjectEmail,
        //       string toEmail, string bodyEmail)
        //{
        //    Outlook.MailItem eMail = (Outlook.MailItem)
        //        this.Application.CreateItem(Outlook.OlItemType.olMailItem);
        //    eMail.Subject = subjectEmail;
        //    eMail.To = toEmail;
        //    eMail.Body = bodyEmail;
        //    eMail.Importance = Outlook.OlImportance.olImportanceLow;
        //    ((Outlook._MailItem)eMail).Send();
        }
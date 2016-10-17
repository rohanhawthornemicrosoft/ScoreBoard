using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Threading;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {

        private TcpListener myListener;
        private IPAddress localaddr;
        private int port = 5050;  // Select any free port you wish
        Thread th;

        delegate void SetTextCallback(string text);

        public Form1()
        {
            InitializeComponent();
        }

        private void log(string msg)
        {
            if (this.messages.InvokeRequired)
            {
                SetTextCallback d = new SetTextCallback(log);
                this.Invoke(d, new object[] { msg });
            }
            else
            {
                this.messages.AppendText(msg + "\n" );
            }
        }

        public string GetTheDefaultFileName(string sLocalDirectory)
        {
            StreamReader sr;
            String sLine = "";

            try
            {
                //Open the default.dat to find out the list
                // of default file
                sr = new StreamReader("data\\Default.Dat");

                while ((sLine = sr.ReadLine()) != null)
                {
                    //Look for the default file in the web server root folder
                    if (File.Exists(sLocalDirectory + sLine) == true)
                        break;
                }
            }
            catch (Exception e)
            {
                log("FAILED: Could not get the default file name");
                Console.WriteLine("An Exception Occurred : " + e.ToString());
            }
            if (File.Exists(sLocalDirectory + sLine) == true)
                return sLine;
            else
                return "";
        }

        public string GetLocalPath(string sMyWebServerRoot, string sDirName)
        {

            StreamReader sr;
            String sLine = "";
            String sVirtualDir = "";
            String sRealDir = "";
            int iStartPos = 0;


            //Remove extra spaces
            sDirName.Trim();

            // Convert to lowercase
            sMyWebServerRoot = sMyWebServerRoot.ToLower();

            // Convert to lowercase
            sDirName = sDirName.ToLower();


            try
            {
                //Open the Vdirs.dat to find out the list virtual directories
                sr = new StreamReader("data\\VDirs.Dat");

                while ((sLine = sr.ReadLine()) != null)
                {
                    //Remove extra Spaces
                    sLine.Trim();

                    if (sLine.Length > 0)
                    {
                        //find the separator
                        iStartPos = sLine.IndexOf(";");

                        // Convert to lowercase
                        sLine = sLine.ToLower();

                        sVirtualDir = sLine.Substring(0, iStartPos);
                        sRealDir = sLine.Substring(iStartPos + 1);

                        if (sVirtualDir == sDirName)
                        {
                            break;
                        }
                    }
                }
            }
            catch (Exception e)
            {
                log("FAILED: Could not get local path");
                Console.WriteLine("An Exception Occurred : " + e.ToString());
            }


            if (sVirtualDir == sDirName)
                return sRealDir;
            else
                return "";
        }

        public string GetMimeType(string sRequestedFile)
        {

            StreamReader sr;
            String sLine = "";
            String sMimeType = "";
            String sFileExt = "";
            String sMimeExt = "";

            // Convert to lowercase
            sRequestedFile = sRequestedFile.ToLower();

            int iStartPos = sRequestedFile.IndexOf(".");

            sFileExt = sRequestedFile.Substring(iStartPos);

            try
            {
                //Open the Vdirs.dat to find out the list virtual directories
                sr = new StreamReader("data\\Mime.Dat");

                while ((sLine = sr.ReadLine()) != null)
                {

                    sLine.Trim();

                    if (sLine.Length > 0)
                    {
                        //find the separator
                        iStartPos = sLine.IndexOf(";");

                        // Convert to lower case
                        sLine = sLine.ToLower();

                        sMimeExt = sLine.Substring(0, iStartPos);
                        sMimeType = sLine.Substring(iStartPos + 1);

                        if (sMimeExt == sFileExt)
                            break;
                    }
                }
            }
            catch (Exception e)
            {
                log("FAILED: Could not get MIME path");
                Console.WriteLine("An Exception Occurred : " + e.ToString());
            }

            if (sMimeExt == sFileExt)
                return sMimeType;
            else
                return "";
        }

        public void SendToBrowser(String sData, ref Socket mySocket)
        {
            SendToBrowser(Encoding.ASCII.GetBytes(sData), ref mySocket);
        }


        public void SendToBrowser(Byte[] bSendData, ref Socket mySocket)
        {
            int numBytes = 0;
            try
            {
                if (mySocket.Connected)
                {
                    if ((numBytes = mySocket.Send(bSendData,
                          bSendData.Length, 0)) == -1)
                        Console.WriteLine("Socket Error cannot Send Packet");
                    else
                    {
                        Console.WriteLine("No. of bytes send {0}", numBytes);
                    }
                }
                else
                    Console.WriteLine("Connection Dropped....");
            }
            catch (Exception e)
            {
                Console.WriteLine("Error Occurred : {0} ", e);
            }
        }

        public void SendHeader(string sHttpVersion, string sMIMEHeader,
            int iTotBytes, string sStatusCode, ref Socket mySocket)
        {

            String sBuffer = "";

            // if Mime type is not provided set default to text/html
            if (sMIMEHeader.Length == 0)
            {
                sMIMEHeader = "text/html";  // Default Mime Type is text/html
            }

            sBuffer = sBuffer + sHttpVersion + sStatusCode + "\r\n";
            sBuffer = sBuffer + "Server: cx1193719-b\r\n";
            sBuffer = sBuffer + "Content-Type: " + sMIMEHeader + "\r\n";
            sBuffer = sBuffer + "Accept-Ranges: bytes\r\n";
            sBuffer = sBuffer + "Content-Length: " + iTotBytes + "\r\n\r\n";

            Byte[] bSendData = Encoding.ASCII.GetBytes(sBuffer);

            SendToBrowser(bSendData, ref mySocket);

            Console.WriteLine("Total Bytes : " + iTotBytes.ToString());

        }

        public void StartListen()
        {

            int iStartPos = 0;
            String sRequest;
            String sDirName;
            String sRequestedFile;
            String sErrorMessage;
            String sLocalDir;
            String sMyWebServerRoot = "C:\\www\\";
            String sPhysicalFilePath = "";
            String sFormattedMessage = "";
            String sResponse = "";


            while (true)
            {
                //Accept a new connection
                Socket mySocket = myListener.AcceptSocket();

                log("Socket Type " + mySocket.SocketType);
                if (mySocket.Connected)
                {
                    log("\nClient Connected!!\n==================\n"+mySocket.RemoteEndPoint) ;


            //make a byte array and receive data from the client 
                    Byte[] bReceive = new Byte[1024];
                    int i = mySocket.Receive(bReceive, bReceive.Length, 0);


                    //Convert Byte to String
                    string sBuffer = Encoding.ASCII.GetString(bReceive);


                    //At present we will only deal with GET type
                    if (sBuffer.Substring(0, 3) != "GET")
                    {
                        log("Only Get Method is supported..");
                        mySocket.Close();
                        return;
                    }


                    // Look for HTTP request
                    iStartPos = sBuffer.IndexOf("HTTP", 1);


                    // Get the HTTP text and version e.g. it will return "HTTP/1.1"
                    string sHttpVersion = sBuffer.Substring(iStartPos, 8);


                    // Extract the Requested Type and Requested file/directory
                    sRequest = sBuffer.Substring(0, iStartPos - 1);


                    //Replace backslash with Forward Slash, if Any
                    sRequest.Replace("\\", "/");


                    //If file name is not supplied add forward slash to indicate 
                    //that it is a directory and then we will look for the 
                    //default file name..
                    if ((sRequest.IndexOf(".") < 1) && (!sRequest.EndsWith("/")))
                    {
                        sRequest = sRequest + "/";
                    }
                    //Extract the requested file name
                    iStartPos = sRequest.LastIndexOf("/") + 1;
                    sRequestedFile = sRequest.Substring(iStartPos);


                    //Extract The directory Name
                    sDirName = sRequest.Substring(sRequest.IndexOf("/"),
                               sRequest.LastIndexOf("/") - 3);

                    /////////////////////////////////////////////////////////////////////
                    // Identify the Physical Directory
                    /////////////////////////////////////////////////////////////////////
                    if (sDirName == "/")
                        sLocalDir = sMyWebServerRoot;
                    else
                    {
                        //Get the Virtual Directory
                        sLocalDir = GetLocalPath(sMyWebServerRoot, sDirName);
                    }


                    log("Directory Requested : " + sLocalDir);

                    //If the physical directory does not exists then
                    // dispaly the error message
                    if (sLocalDir.Length == 0)
                    {
                        sErrorMessage = "<H2>Error!! Requested Directory does not exists</H2><Br>";
                        //sErrorMessage = sErrorMessage + "Please check data\\Vdirs.Dat";

                        //Format The Message
                        SendHeader(sHttpVersion, "", sErrorMessage.Length," 404 Not Found", ref mySocket);

                        //Send to the browser
                        SendToBrowser(sErrorMessage, ref mySocket);

                        mySocket.Close();

                        continue;
                    }

                    /////////////////////////////////////////////////////////////////////
                    // Identify the File Name
                    /////////////////////////////////////////////////////////////////////

                    //If The file name is not supplied then look in the default file list
                    if (sRequestedFile.Length == 0)
                    {
                        // Get the default filename
                        sRequestedFile = GetTheDefaultFileName(sLocalDir);

                        if (sRequestedFile == "")
                        {
                            sErrorMessage = "<H2>Error!! No Default File Name Specified</H2>";
                            SendHeader(sHttpVersion, "", sErrorMessage.Length, " 404 Not Found", ref mySocket);
                            SendToBrowser(sErrorMessage, ref mySocket);

                            mySocket.Close();

                            return;

                        }
                    }

                    //////////////////////////////////////////////////
                    // Get TheMime Type
                    //////////////////////////////////////////////////

                    String sMimeType = GetMimeType(sRequestedFile);


                    //Build the physical path
                    sPhysicalFilePath = sLocalDir + sRequestedFile;
                    Console.WriteLine("File Requested : " + sPhysicalFilePath);

                    if (File.Exists(sPhysicalFilePath) == false)
                    {

                        sErrorMessage = "<H2>404 Error! File Does Not Exists...</H2>";
                        SendHeader(sHttpVersion, "", sErrorMessage.Length, " 404 Not Found", ref mySocket);
                        SendToBrowser(sErrorMessage, ref mySocket);

                        log(sFormattedMessage);
                        Console.WriteLine(sFormattedMessage);
                    }
                    else
                    {
                        int iTotBytes = 0;

                        sResponse = "";

                        FileStream fs = new FileStream(sPhysicalFilePath, FileMode.Open, FileAccess.Read, FileShare.Read);
                        // Create a reader that can read bytes from the FileStream.


                        BinaryReader reader = new BinaryReader(fs);
                        byte[] bytes = new byte[fs.Length];
                        int read;
                        while ((read = reader.Read(bytes, 0, bytes.Length)) != 0)
                        {
                            // Read from the file and write the data to the network
                            sResponse = sResponse + Encoding.ASCII.GetString(bytes, 0, read);
                            iTotBytes = iTotBytes + read;
                        }
                        reader.Close();
                        fs.Close();

                        SendHeader(sHttpVersion, sMimeType, iTotBytes, " 200 OK", ref mySocket);
                        SendToBrowser(bytes, ref mySocket);
                        //mySocket.Send(bytes, bytes.Length,0);

                    }
                    mySocket.Close();
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
                try
                {
                //start listing on the given port
                IPAddress.TryParse("127.0.0.1", out localaddr);
                myListener = new TcpListener(localaddr, port);
                myListener.Start();
                log("Web Server Running... Press STOP to Cancel...");
                button2.Show();

                //start the thread which calls the method 'StartListen'
                th = new Thread(new ThreadStart(StartListen));
                th.Start();

                }
                catch (Exception ex)
                {
                log("FAILED to open listener...");
                Console.WriteLine("An Exception Occurred while Listening :" + ex.ToString());
                }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                myListener.Stop();
                th.Abort();
                log("ABORTED");
            } catch (Exception ex)
            {
                Console.WriteLine("An Exception Occurred while aborting :" + ex.ToString());
            }
            button2.Hide();
        }
    }
}

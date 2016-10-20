using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Threading;
using Microsoft.Office.Interop.Excel;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        private TcpListener myListener;
        private IPAddress localaddr;
        private int port = 5050;  // Select any free port you wish
        Thread th;
        string spreadsheetfile = "data\\Spreadsheet.Dat";
        string passwordfile = "data\\Password.Dat";
        int results = 0;
        Dictionary<string, string> result = new Dictionary<string, string>();
        const int post = 4;
        const int get = 3;
        delegate void SetTextCallback(string text);
        Microsoft.Office.Interop.Excel.Application excelApp = null;
        Workbooks wkbks = null;
        Workbook wkbk = null;

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
                    if ((numBytes = mySocket.Send(bSendData, bSendData.Length, 0)) == -1) log("Socket Error cannot Send Packet");
                }
                else
                    log("Connection Dropped....");
            }
            catch (Exception e)
            {
                log("Error Occurred: " + e.ToString());
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
        }

        Boolean parse_login(Dictionary<string, string> postpairs)
        {
            if (File.Exists(passwordfile))
            {
                string password = System.IO.File.ReadAllText(passwordfile);
                log("testing entered password - " + postpairs["pass"] + " - against stored password - " + password);
                if (postpairs["pass"].Equals(password)) return true; else return false;
            }
            log("Password not set");
            return false;
        }

        string calculate_result(Dictionary<string, string> postpairs)
        {
            string position;
            position = postpairs.ElementAt(postpairs.Count - 2).Key;
            log("got a result from " + position);
            if (!result.ContainsKey(position))
            {
                result[position] = postpairs[position];
                results++;
            }
            return position + ".dpi";
        }

        void get_spreadsheet()
        {
            string spreadsheet;
            if (File.Exists(spreadsheetfile)) spreadsheet = System.IO.File.ReadAllText(spreadsheetfile);
            else { log("Spreadsheet not selected"); return; }

            bool wasFoundRunning = false;

            Microsoft.Office.Interop.Excel.Application tApp = null;
            //Checks to see if excel is opened
            try
            {
                tApp = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                wasFoundRunning = true;
            }
            catch (Exception)//Excel not open
            {
                wasFoundRunning = false;
                log("Excel not open");
                return;
            }
            finally
            {
                if (true == wasFoundRunning)
                {
                    excelApp = tApp;
                    log("Excel was running... Using active spreadsheet");
                    wkbk = excelApp.ActiveWorkbook;
                }
                else
                {
                    log("Excel was not running... Opening spreadsheet");
                    excelApp = new Microsoft.Office.Interop.Excel.Application();
                    wkbks = excelApp.Workbooks;
                    wkbk = wkbks.Open(spreadsheet);
                }
                //Release the temp if in use
                //if (null != tApp) { System.Runtime.InteropServices.Marshal.FinalReleaseComObject(tApp); }
                //tApp = null;
            }

        }

        void press_button(int i)
        {
            string macro = null;
            if (excelApp == null)
            {
                log("Cannot access workbook");
                return;
            }
            if (i > 1) macro = "Good"; else macro = "NoLift";
            try
            {
                excelApp.Run(macro);
            }
            catch (Exception e)
            {
                log("Can't run " + macro + " because: " + e.ToString());
            }
        }

        Dictionary<string, string> show_results()
        {
            Dictionary<string, string> screen = new Dictionary<string, string>();
            int i = 0;
            if (results == 3)
            {
                foreach (KeyValuePair<string, string> kvp in result)
                {
                    screen.Add(kvp.Key, kvp.Value);
                    if (kvp.Value.Equals("good")) i++;
                }
                press_button(i);
                result.Clear();
                results = 0;
                return screen;
            }
            screen.Add("left", "blank");
            screen.Add("head", "blank");
            screen.Add("right", "blank");
            return screen;
        }

        public void StartListen()
        {

            int iStartPos = 0;
            String sRequest;
            String sDirName;
            String sRequestedFile;
            String sErrorMessage;
            String sLocalDir;
            String sMyWebServerRoot = "C:\\scoreboard\\";
            String sPhysicalFilePath = "";
            String sResponse = "";

            get_spreadsheet();

            while (true)
            {
                //Accept a new connection
                Socket mySocket = myListener.AcceptSocket();

                if (mySocket.Connected)
                {

                    //make a byte array and receive data from the client 
                    Byte[] bReceive = new Byte[1024];
                    int i = mySocket.Receive(bReceive, bReceive.Length, SocketFlags.None);
                    int headlength;
                    Dictionary<string, string> postpairs = new Dictionary<string, string>();
                            
                    //Convert Byte to String
                    string sBuffer = Encoding.ASCII.GetString(bReceive);

                    //At present we will only deal with GET type
                    if (sBuffer.Substring(0, 3) != "GET" && sBuffer.Substring(0, 4) != "POST")
                    {
                        log("Only Get and Post Method is supported, instead I got this:");
                        log(sBuffer);
                        mySocket.Close();
                        return;
                    }

                    if (sBuffer.Substring(0, 4) == "POST")
                    {
                        headlength = post;

                        // create dictionary of POST pairs
                        int start = sBuffer.Length - 1;
                        while (sBuffer[start] != '\n') start--; start++;
                        string[] pairs = sBuffer.Substring(start).Split('&');
                        foreach (string pairfound in pairs)
                        {
                            string[] splits = pairfound.Split('=');
                            postpairs.Add(splits[0], splits[1]);
                        }

                        foreach (KeyValuePair<string, string> kvp in postpairs) log("got pair: " + kvp.Key + " = " + kvp.Value);     
                    }
                    else headlength = get;

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
                    sDirName = sRequest.Substring(sRequest.IndexOf("/"), sRequest.LastIndexOf("/") - headlength);

                    // do something fancy if the requested file is important
                    if (sRequestedFile.Equals("login.dpi"))
                    {
                        if (parse_login(postpairs)) sRequestedFile = "choose.dpi";
                    }
                    if (sRequestedFile.Equals("display.dpi"))
                    {
                        sRequestedFile = postpairs["choice"] + ".dpi";
                        log(postpairs["choice"] + " has joined");
                    }
                    if (sRequestedFile.Equals("judge.dpi"))
                    {
                        sRequestedFile = calculate_result(postpairs);
                    }
                    if (sRequestedFile.Equals("screen.dpi"))
                    {
                        postpairs = show_results();
                        headlength = post;
                    }


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

                    if (File.Exists(sPhysicalFilePath) == false)
                    {

                        sErrorMessage = "<H2>404 Error! File Does Not Exists...</H2>";
                        SendHeader(sHttpVersion, "", sErrorMessage.Length, " 404 Not Found", ref mySocket);
                        SendToBrowser(sErrorMessage, ref mySocket);
                    }
                    else
                    {
                        int iTotBytes = 0;

                        sResponse = "";

                        // open the file
                        FileStream fs = new FileStream(sPhysicalFilePath, FileMode.Open, FileAccess.Read, FileShare.Read);

                        // Create a reader that can read bytes from the FileStream
                        BinaryReader reader = new BinaryReader(fs);

                        byte[] bytes = new byte[fs.Length];
                        int read;

                        // Read from the file and write the data to the network
                        while ((read = reader.Read(bytes, 0, bytes.Length)) != 0)
                        {
                            sResponse += Encoding.ASCII.GetString(bytes, 0, read);
                            iTotBytes += read;
                        }
                        reader.Close();
                        fs.Close();

                        // parse the file if post
                        if (headlength == post)
                        {
                            string body = sResponse;
                            foreach (KeyValuePair<string, string> kvp in postpairs) if (kvp.Value.Length > 0) body = body.Replace("#" + kvp.Key + "#", kvp.Value);
                            byte[] newbytes = new byte[body.Length * sizeof(char)];
                            newbytes = Encoding.ASCII.GetBytes(body);
                            bytes = newbytes;
                            iTotBytes = newbytes.Length;
                        }

                        SendHeader(sHttpVersion, sMimeType, iTotBytes, " 200 OK", ref mySocket);
                        SendToBrowser(bytes, ref mySocket);

                    }
                    mySocket.Close();
                }
            }
        }

        private IPAddress LocalIPAddress()
        {
            if (!System.Net.NetworkInformation.NetworkInterface.GetIsNetworkAvailable())
            {
                return null;
            }

            IPHostEntry host = Dns.GetHostEntry(Dns.GetHostName());

            return host
                .AddressList
                .FirstOrDefault(ip => ip.AddressFamily == AddressFamily.InterNetwork);
        }

        private void button1_Click(object sender, EventArgs e)
        {
                try
                {
                //start listing on the given port on the local machine
                localaddr = LocalIPAddress();
                if (localaddr == null)
                {
                    log("Not connected to a network");
                    return;
                }
                myListener = new TcpListener(localaddr, port);
                myListener.Start();
                log("Web Server Running at http://"+localaddr.ToString()+":"+port.ToString()+" Press STOP to Cancel...");
                button2.Show();

                //start the thread which calls the method 'StartListen'
                th = new Thread(new ThreadStart(StartListen));
                th.Start();

                }
                catch (Exception ex)
                {
                log("FAILED to open listener : " + ex.ToString());
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

        private void button3_Click(object sender, EventArgs e)
        {
            OpenFileDialog fdlg = new OpenFileDialog();
            fdlg.Title = "Select Next Lifter Spreadsheet";
            fdlg.InitialDirectory = @"c:\";
            fdlg.Filter = "All files (*.*)|*.*|All files (*.*)|*.*";
            fdlg.FilterIndex = 2;
            fdlg.RestoreDirectory = true;
            if (fdlg.ShowDialog() == DialogResult.OK)
            {
                string spreadsheet = fdlg.FileName;
                if (File.Exists(spreadsheetfile)) File.Delete(spreadsheetfile);
                System.IO.File.WriteAllText(spreadsheetfile, spreadsheet);
                log("Selected " + spreadsheet);
            }
        }

        private void buttonSetPassword_Click(object sender, EventArgs e)
        {
            string password;
            password = textPassword.Text;
            if (File.Exists(passwordfile)) File.Delete(passwordfile);
            System.IO.File.WriteAllText(passwordfile, password);
            log("Password set to " + password);
        }
    }
}

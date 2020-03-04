// text2pdf_20.cs - .NET DLL wrapper for FyTek's Text2PDF

// Call this DLL from the language of your choice as long as it supports
// a COM or .NET DLL.  May be 32 or 64-bit.
// This DLL accepts parameters and then builds the PDF which you may save
// locally, on the box the server is running on (if it's different box) or
// have the PDF returned as a byte array for display on website or for saving
// in a database.
// This DLL calls the Text2PDF executable (32 or 64-bit) or uses sockets
// when Text2PDF is running as a server.  It sends the parameter settings
// made here to build the PDF.  For exapmle, you might want to startup Text2PDF
// with a pool of 5 connections on a Linux box and call it from this DLL on a
// Windows box.  Even if you run Text2PDF on the same box it's recommended
// to start a Text2PDF server in order to keep resource usage in check.
// Note you may start more than one Text2PDF server at a time with different
// port numbers for each one.
// Use startServer to start up a Text2PDF server and stopServer to shut it down.
// You probably want to do that outside of your main routine that will be building
// PDFs as your main routine will link to this DLL to call the already running
// service.

// Many of the methods have a legacy method underneath.  These are to support
// the names used by the older dll.

using System.Diagnostics;
using System.Collections.Generic;
using System.IO;
using System;
using System.Net;
using System.Net.Sockets;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Security.Cryptography;
using System.Globalization;

// Compiling this code:
// Microsoft.Net.Compilers.3.4.0\tools\csc /target:library /platform:anycpu /out:text2pdf_20.dll text2pdf_20.cs /keyfile:mykey.snk
// C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe /codebase c:text2pdf_20.dll
// cscript.exe (or wscript.exe) rwtest.vbs

namespace FyTek
{
    [ComVisible(true)]
    [Guid("821E5678-328C-48FA-B9E3-A33D53C475E7")]
	[ProgId("FyTek.Text2PDF")]    
    public class Text2PDF : IDisposable
    {
        void IDisposable.Dispose()
        {

        }
        private TcpClient client = new TcpClient();
        private NetworkStream stream;
        private String exe = "text2pdf64"; // the executable - change with setExe
        private const String srvHost = "localhost";
        private const int srvPort = 7080;
        private const int srvPool = 5;
        private static String srvFile = ""; // the file of servers and ports
        private static int srvNum = 0; // the array index for the next server to use
        private bool useAvailSrv = false; // true when choosing the next available server
        private List<string> cmds = new List<string>(); // input commands (input.txt)
        private Dictionary<string,object> opts = new Dictionary<string,object>(); // all of the parameter settings from the method calls
        private Dictionary<string,object> server = new Dictionary<string,object>(); // the server host/port/log file key/values

        private String units = "";
        private Double unitsMult = 1;

        [ComVisible(true)]
        public class Results {
            public byte[] Bytes {get; set;}
            public String Msg {get; set;}
            public int Pages {get; set;}
        }        

        private class Server {
            public String Host {get; set;}
            public int Port {get; set;}
            public Server(String host, int port){
                this.Host = host;
                this.Port = port;
            }
        }

        private static List<Server> servers = new List<Server>();

        // Start up Text2PDF as a server
        [ComVisible(true)]
        public String setServerFile(String fileName){
            srvFile = fileName;
            String line = "";
            String[] retCmds = new String[2]; 
            Regex r = new Regex("[\\s\\t]+");
            int port;
            srvNum = 0;
            try {
                System.IO.StreamReader file =   
                new System.IO.StreamReader(fileName);  
                servers = new List<Server>();
                while((line = file.ReadLine()) != null)  
                {  
                    if (!line.Trim().StartsWith("#")
                    && !line.Trim().Equals("")){
                        retCmds = r.Split(line.Trim());
                        if (retCmds[0].Equals("exe")){
                            setExe(retCmds[1]); // passing location of exe instead of a host/port
                        } else {
                            int.TryParse(retCmds[1],out port);
                            if (port > 0)
                                servers.Add(new Server(retCmds[0],port));
                        }
                    }
                }  
                file.Close();                              
            } catch (IOException e) {
                return e.Message;
            }
            return "";
        }

        // Start up Text2PDF as a server
        [ComVisible(true)]
        public String startServer(
            String host = srvHost,
            int port = srvPort,
            int pool = srvPool,
            String log = ""
          )
       {
            byte[] bytes = {};
            String errMsg = "";            
            server["host"] = host;
            server["port"] = port;
            server["pool"] = pool;
            server["log"] = log; // file on server to log the output
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.CreateNoWindow = false;
            startInfo.UseShellExecute = true;
            startInfo.FileName = exe;
            startInfo.WindowStyle = ProcessWindowStyle.Hidden;
            String cmdsOut = "";
            object s = "";

            if (opts.TryGetValue("licname", out s)){
                object p = "";
                object d = false;
                opts.TryGetValue("licpwd", out p);
                opts.TryGetValue("licweb", out d);
                cmdsOut += " -licname " + s + " -licpwd " + p + ((bool) d ? " -licweb" : "");
            } else {
                if (!opts.TryGetValue("kn", out s)){
                    setKeyName("demo");
                }
            }
            if (opts.TryGetValue("kc", out s))
                cmdsOut += " -kc " + s;
            if (opts.TryGetValue("kn", out s))
                cmdsOut += " -kn " + s;

            cmdsOut += (!log.Equals("") ? " -log " + '"' + log + '"' : "")
                + " -port " + port + " -pool " + pool + " -host " + host;
            startInfo.Arguments = "-server " + cmdsOut;            
            (bytes, errMsg) = runProcess(startInfo, false);
            opts.Remove("licname");
            opts.Remove("licpwd");
            opts.Remove("licweb");
            opts.Remove("kn");
            opts.Remove("kc");
            return errMsg;
        }

        [ComVisible(true)]
        public void setServer(
            String host = srvHost,
            int port = srvPort
          )
       {
            server["host"] = host;
            server["port"] = port;
       }

        // Stop the server
        [ComVisible(true)]
        public String stopServer(){
            byte[] bytes;
            String msg;
            int pages;
            setOpt("serverCmd","-quit");
            (bytes, msg, pages) = callTCP(isStop: true);
            return msg;
        }

        // Get the current stats for the server
        [ComVisible(true)]
        public String serverStatus(bool allServers = false){
            byte[] bytes;
            String msg = "";
            String sMsg = "";
            int pages = 0;
            if (!allServers || servers.Count == 0){
                if (!isServerRunning()){
                    return "Server " + server["host"] + " is not responding on port " + server["port"] + ".";
                }
                setOpt("serverCmd","-serverstat");            
                (bytes, msg, pages) = callTCP(isStatus: true);            
            } else {
                foreach(var item in servers){
                    server["host"] = item.Host;
                    server["port"] = item.Port;
                    setOpt("serverCmd","-serverstat");            
                    (bytes, sMsg, pages) = callTCP(isStatus: true);
                    msg += sMsg + "\n";
                }
            }
            return msg;
        }

        // Get the current stats for the server
        [ComVisible(true)]
        public int serverThreads(){
            byte[] bytes;
            String msg;
            int pages = 0;
            int threadsAvail = 0;
            if (!isServerRunning()){
                return 0;
            }
            setOpt("serverCmd","-threadsavail");
            (bytes, msg, pages) = callTCP(isStatus: true);
            int.TryParse(msg, out threadsAvail);
            return threadsAvail;
        }

        // Stop a process
        [ComVisible(true)]
        public String serverCancelId(int id){
            byte[] bytes;
            String msg;
            int pages = 0;
            setOpt("stopId","-stopid " + id);
            (bytes, msg, pages) = callTCP();
            return msg;
        }

        // Check server
        [ComVisible(true)]
        public bool isServerRunning(){
            Object host = "";
            int tryCount = 0;
            bool srvRunning = false;
            if (client.Connected){
                return true;
            }            
            if (!server.TryGetValue("host", out host) && servers.Count == 0){
                setServer();
            }
            if (host == null && servers.Count > 0){
                // Loop through list of servers and get the next one that gets connected
                while(tryCount < servers.Count){
                    srvNum++;
                    srvNum %= servers.Count;
                    useAvailSrv = true;
                    srvRunning = true;
                    setServer(servers[srvNum].Host, servers[srvNum].Port);
                    if (client.Connected){
                        // if previous connection, disconnect it
                        try {
                            stream.Close();
                            client.Close();                        
                        } catch (SocketException) { }
                    }
                    try { 
                        client = new TcpClient((String) server["host"], (int) server["port"]);                        
                        stream = client.GetStream();
                        Socket s = client.Client;
                        srvRunning = !((s.Poll(1000, SelectMode.SelectRead) && (s.Available == 0)) || !s.Connected);
                        return true;
                    } catch (SocketException) {
                        srvRunning = false;
                        stream.Close();
                        client.Close();                        
                    }
                    tryCount++;
                }
            }            
            if (!client.Connected && server.TryGetValue("host", out host)){
                try {
                    client = new TcpClient((String) server["host"], (int) server["port"]);
                    stream = client.GetStream();                    
                    Socket s = client.Client;
                    return !((s.Poll(1000, SelectMode.SelectRead) && (s.Available == 0)) || !s.Connected);
                } catch (SocketException) {
                    return false;
                }
            }
            return false;
        }

        // Build the PDF using the server, optionally return the 
        // raw bytes of the PDF for further processing when retBytes = true
        // optional file name as well if saveFile is passed - this allows
        // for saving file on this box if server is running on different box
        private object buildPDFTCP(bool retBytes = false, String saveFile = ""){            
            byte[] bytes = {};
            String errMsg = "";
            int numPages = 0;
            object s = "";

            if (cmds.Count > 0){
                if (!opts.TryGetValue("inFile", out s)){
                    s = $@"{Guid.NewGuid()}.txt"; // come up with unique file name
                    setInFile((String) s);
                    sendFileTCP((String) s, "--input--commands--");
                }
            }

            if (!saveFile.Equals("") || retBytes){
                if (!opts.TryGetValue("outFile", out s)){
                    setOutFile("membuild");
                }
            }

            (bytes, errMsg, numPages) = callTCP(retBytes: retBytes, saveFile: saveFile);
            Results ret = new Results();
            ret.Bytes = bytes;
            ret.Msg = errMsg.Equals("") ? "OK" : errMsg;
            ret.Pages = numPages;
            if (useAvailSrv){
                server.Clear();
                useAvailSrv = false;
            }
            return ret;
        }

        // Send all files to server - only necessary if server is on a different box
        [ComVisible(true)]
        public void setAutoSendFiles(){
          setOpt("autosend",true);
        }

        // Pass file contents in memory for input type files
        [ComVisible(true)]
        public String addText(String a)
        {
            cmds.Add(a);
            return a;
        }
        [Obsolete("PDFCmd is deprecated, please use addText instead.")]
        [ComVisible(true)]
        public String PDFCmd(String a){
            return addText(a);
        }

        // Pass file contents in memory for input type files
        [ComVisible(true)]
        public int setAutoSize(int lines = 0)
        {
            if (lines == 0){
                setOpt("autosize",true);
            } else  {
                setOpt("autosize",lines);
            }
            return lines;
        }
        [Obsolete("AutoSize is deprecated, please use setAutoSize instead.")]
        [ComVisible(true)]
        public int AutoSize(int lines = 0){
            return setAutoSize(lines);
        }

        // Assign the key name
        [ComVisible(true)]
        public String setKeyName(String a)
        {
            if (a.ToLower().Equals("demo")){
                // Get the demo key from website - this only works with the demo text2pdf executable
                WebClient wClient = new WebClient();
                string res = wClient.DownloadString("http://www.fytek.com/cgi-bin/genkeyw_v2.cgi?prod=text2pdf");
                Regex regex = new Regex("-kc [A-Z0-9]*");
                Match match = regex.Match(res);                
                setOpt("kn","testkey");
                setOpt("kc",match.Value.Substring(4));
            } else {
                setOpt("kn",a);
            }
            return a;
        }

        // Assign the key code
        [ComVisible(true)]
        public String setKeyCode(String a)
        {
            setOpt("kc",a);
            return a;
        }

        // License settings
        [ComVisible(true)]
        public void licInfo(String licName,
            String licPwd,
            int autoDownload)
        {
            setOpt("licname",licName);
            setOpt("licpwd",licPwd);
            if (autoDownload == 1){
                setOpt("licweb",true);
            }
        }

        [ComVisible(true)]
        public void licInfo(String licName,
            String licPwd,
            bool autoDownload)
        {
            setOpt("licname",licName);
            setOpt("licpwd",licPwd);
            if (autoDownload){
                setOpt("licweb",true);
            }
        }

        // Assign the user
        [ComVisible(true)]
        public String setUser(String a)
        {
            setOpt("u",a);
            return a;
        }
        [Obsolete("User is deprecated, please use setUser instead.")]
        [ComVisible(true)]
        public String User(String a) {return setUser(a);}
            
        // Assign the owner
        [ComVisible(true)]
        public String setOwner(String a)
        {
            setOpt("o",a);
            return a;
        }
        [Obsolete("Owner is deprecated, please use setOwner instead.")]
        public String Owner(String a) {return setOwner(a);}

        // Set no annote
        [ComVisible(true)]
        public void setNoAnnote()
        {
            setOpt("noannote",true);
        }
        [Obsolete("NoAnnote is deprecated, please use setNoAnnote instead.")]
        [ComVisible(true)]
        public void NoAnnote(){setNoAnnote();}

        // Set no copy
        [ComVisible(true)]
        public void setNoCopy()
        {
            setOpt("nocopy",true);
        }
        [Obsolete("NoCopy is deprecated, please use setNoCopy instead.")]
        [ComVisible(true)]
        public void NoCopy(){setNoCopy();}

        // Set no change
        [ComVisible(true)]
        public void setNoChange()
        {
            setOpt("nochange",true);
        }
        [Obsolete("NoChange is deprecated, please use setNoChange instead.")]
        [ComVisible(true)]
        public void NoChange(){setNoChange();}

        // Set no print
        [ComVisible(true)]
        public void setNoPrint()
        {
            setOpt("noprint",true);
        }
        [Obsolete("NoPrint is deprecated, please use setNoPrint instead.")]
        [ComVisible(true)]
        public void NoPrint(){setNoPrint();}

        // Set no rights
        [ComVisible(true)]
        public void setNoRights()
        {
            setOpt("norights",true);
        }
        [Obsolete("NoRights is deprecated, please use setNoRights instead.")]
        [ComVisible(true)]
        public void NoRights(){setNoRights();}

        // Set the GUI process window off
        [ComVisible(true)]
        public void setGUIOff()
        {
            setOpt("guioff",true);
        }

        // Set any other command line type options
        [ComVisible(true)]
        public String setCmdlineOpts(String a)
        {
            object s = "";
            if (opts.TryGetValue("extOpts", out s))
                a = s + " " + a;
            setOpt("extOpts",a);
            return a;
        }

        // Assign the executable
        [ComVisible(true)]
        public String setExe(String a)
        {
            exe = a;
            return a;
        }

        // Assign the input file name
        [ComVisible(true)]
        public String setInFile(String a)
        {
            setOpt("inFile",a);
            return a;
        }
        [Obsolete("InputFile is deprecated, please use setInFile instead.")]
        [ComVisible(true)]
        public String InputFile(String a){return setInFile(a);}

        // Assign the output file name
        [ComVisible(true)]
        public String setOutFile(String a)
        {
            setOpt("outFile",a);
            return a;
        }
        [Obsolete("OutputFile is deprecated, please use setOutFile instead.")]
        [ComVisible(true)]
        public String OutputFile(String a){return setOutFile(a);}

        // Assign the output file name
        [ComVisible(true)]
        public String setFileMask(String a)
        {
            setOpt("mask",a);
            return a;
        }
        [Obsolete("FileMask is deprecated, please use setFileMask instead.")]
        [ComVisible(true)]
        public String FileMask(String a){return setFileMask(a);}
        
        [ComVisible(true)]
        public void setSubDir()
        {
            setOpt("s",true);
        }
        [Obsolete("SubDir is deprecated, please use setSubDir instead.")]
        [ComVisible(true)]
        public void SubDir(){setSubDir();}

        [ComVisible(true)]
        public void setMail()
        {
            setOpt("mail",true);
        }
        [Obsolete("AutoMail is deprecated, please use setMail instead.")]
        [ComVisible(true)]
        public void AutoMail(){setMail();}

        // Optimize
        [ComVisible(true)]
        public void setOptimize(bool compress = true)
        {
            if (compress){
                setOpt("opt",true);
            } else {
                setOpt("opt15",true);
            }            
        }
        [Obsolete("Optimize is deprecated, please use setOptimize instead.")]
        [ComVisible(true)]
        public void Optimize(){setOptimize();}

        // Compression 1.5
        [ComVisible(true)]
        public void setComp15()
        {
            setOpt("comp15",true);
        }
        [Obsolete("Comp15 is deprecated, please use setComp15 instead.")]
        [ComVisible(true)]
        public void Comp15(){setComp15();}

        // Encrypt 128
        [ComVisible(true)]
        public void setEncrypt128()
        {
            setOpt("e128",true);
        }
        [Obsolete("Encrypt128 is deprecated, please use setEncrypt128 instead.")]
        [ComVisible(true)]
        public void Encrypt128(){setEncrypt128();}

        // AES 128
        [ComVisible(true)]
        public String setEncryptAES(String a)
        {
            setOpt("aes",a);
            return a;
        }
        [Obsolete("EncryptAES is deprecated, please use setEncryptAES instead.")]
        [ComVisible(true)]
        public String EncryptAES(String a){return setEncryptAES(a);}

        // open output
        [ComVisible(true)]
        public void setForwardRef()
        {
            setOpt("forwardref",true);
        }        
        [Obsolete("ForwardRef is deprecated, please use setForwardRef instead.")]
        [ComVisible(true)]
        public void ForwardRef(){setForwardRef();}


        // overwrite existing
        [ComVisible(true)]
        public void setForce()
        {
            setOpt("force",true);
        }

        // open output
        [ComVisible(true)]
        public void setOpen()
        {
            setOpt("open",true);
        }        
        [Obsolete("AutoOpen is deprecated, please use setOpen instead.")]
        [ComVisible(true)]
        public void AutoOpen(){setOpen();}

        // print output
        [ComVisible(true)]
        public void setPrint()
        {
            setOpt("print",true);
        }       
        [Obsolete("AutoPrint is deprecated, please use setPrint instead.")]
        [ComVisible(true)]
        public void AutoPrint(){setPrint();}

        // buildlog file
        [ComVisible(true)]
        public String setBuildLog(String fileName)
        {
            setOpt("buildlog",fileName);
            return fileName;
        }      

        // debug
        [ComVisible(true)]
        public String setDebug(String fileName)
        {
            setOpt("debug",fileName);
            return fileName;
        }        

        // errFile
        [ComVisible(true)]
        public String setErrFile(String fileName)
        {
            setOpt("e",fileName);
            return fileName;
        }        

        [ComVisible(true)]
        public String setWorkingDir(String a)
        {
            setOpt("cwd",a);
            return a;
        }        
        [Obsolete("WorkingDir is deprecated, please use setWorkingDir instead.")]
        [ComVisible(true)]
        public String WorkingDir(String a){return setWorkingDir(a);}


        // subject
        [ComVisible(true)]
        public String setSubject(String a)
        {
            setOpt("subject",a);
            return a;
        }        
        [Obsolete("DocSubject is deprecated, please use setSubject instead.")]
        [ComVisible(true)]
        public String DocSubject(String a){return setSubject(a);}

        // author
        [ComVisible(true)]
        public String setAuthor(String a)
        {
            setOpt("author",a);
            return a;
        }        
        [Obsolete("DocAuthor is deprecated, please use setAuthor instead.")]
        [ComVisible(true)]
        public String DocAuthor(String a){return setAuthor(a);}

        // keywords
        [ComVisible(true)]
        public String setKeywords(String a)
        {
            setOpt("keywords",a);
            return a;
        }        
        [Obsolete("DocKeywords is deprecated, please use setKeywords instead.")]
        [ComVisible(true)]
        public String DocKeywords(String a){return setKeywords(a);}

        // creator
        [ComVisible(true)]
        public String setCreator(String a)
        {
            setOpt("creator",a);
            return a;
        }        
        [Obsolete("DocCreator is deprecated, please use setCreator instead.")]
        [ComVisible(true)]
        public String DocCreator(String a){return setCreator(a);}

        // producer
        [ComVisible(true)]
        public String setProducer(String a)
        {
            setOpt("producer",a);
            return a;
        }      
        [Obsolete("DocProducer is deprecated, please use setProducer instead.")]
        [ComVisible(true)]
        public String DocProducer(String a){return setProducer(a);}
        
        [ComVisible(true)]
        public String setCreationDate(String a = "")
        {
            setOpt("creationdate",a);
            return a;
        }      

        [ComVisible(true)]
        public String setModDate(String a = "")
        {
            setOpt("moddate",a);
            return a;
        }      

        [ComVisible(true)]
        public String setUnits(String a)
        {
          units = a.ToLower();
          unitsMult *= 72;
          switch (units) {
            case "cm":
              unitsMult = 72 / 2.54;
              break;
            case "mm":
              unitsMult = 72 / 25.4;
              break;
            case "in":
              unitsMult = 72;
              break;
            default:
              unitsMult = 1;
              break;
          }
          unitsMult /= 72;
          return a;
        }      
        [Obsolete("Units is deprecated, please use setUnits instead.")]
        [ComVisible(true)]
        public String Units(String a){return setUnits(a);}    

        [ComVisible(true)]
        public String setBkgPDF(String a, String passwd = ""){
            setOpt("pdf",a);
            setOpt("bkgpass",passwd);
            return a;            
        }
        [Obsolete("BkgPDF is deprecated, please use setBkgPDF instead.")]
        [ComVisible(true)]
        public String BkgPDF(String a){return setBkgPDF(a, "");}

        [ComVisible(true)]
        public String setPageSize(Double width, Double height)
        {
          String a = width * unitsMult + "," + height * unitsMult;
          setOpt("pagesize",a);
          return a;
        }      
        [Obsolete("PageSize is deprecated, please use setPageSize instead.")]
        [ComVisible(true)]
        public String PageSize(Double width, Double height){return setPageSize(width, height);}

        [ComVisible(true)]
        public Double setFontSize(Double a)
        {
          setOpt("fontsize",a.ToString());
          return a;
        }      

        [ComVisible(true)]
        public Double setPageScale(Double x)
        {
            setOpt("scale",x.ToString());
          return x;
        }       
        [Obsolete("PageScale is deprecated, please use setPageScale instead.")]
        [ComVisible(true)]
        public Double PageScale(Double x){return setPageScale(x);}

        [ComVisible(true)]
        public Double setPageRight(Double x)
        {
          setOpt("right",(x * unitsMult).ToString());          
          return x;
        }       
        [Obsolete("PageRight is deprecated, please use setPageRight instead.")]
        [ComVisible(true)]
        public Double PageRight(Double a){return setPageRight(a);}

        [ComVisible(true)]
        public Double setPageDown(Double x)
        {
          setOpt("down",(x * unitsMult).ToString());          
          return x;
        }  
        [Obsolete("PageDown is deprecated, please use setPageDown instead.")]
        [ComVisible(true)]
        public Double PageDown(Double a){return setPageDown(a);}

        [ComVisible(true)]
        public Double setPointPct(Double x)
        {
          setOpt("pointpct",x);          
          return x;
        }  
        [Obsolete("PointPct is deprecated, please use setPointPct instead.")]
        [ComVisible(true)]
        public Double PointPct(Double a){return setPointPct(a);}

        [ComVisible(true)]
        public Double setTextCompress(Double x)
        {
          setOpt("comp",x);          
          return x;
        }  
        [Obsolete("TextCompress is deprecated, please use setTextCompress instead.")]
        [ComVisible(true)]
        public Double TextCompress(Double a){return setTextCompress(a);}

        [ComVisible(true)]        
        public String setBkgImg(String fileName,
            Double x,
            Double y,
            Double scalex = double.MinValue,
            Double scaley = double.MinValue)
        {          
            setOpt("img",fileName + "," + x + "," + y + (scalex != double.MinValue && scaley != double.MinValue ? "," + scalex + "," + scaley : ""));
            return fileName;
        }      
        [Obsolete("BkgImg is deprecated, please use setBkgImg instead.")]
        [ComVisible(true)]
        public String BkgImg(String fileName,
            Double x,
            Double y,
            Double scalex = double.MinValue,
            Double scaley = double.MinValue){return setBkgImg(fileName, x, y, scalex, scaley);}

        [ComVisible(true)]
        public int setBkgPage(int pgnum)
        {
            if (pgnum > 0){
                setOpt("pdfpage",pgnum + "");          
            }
            return pgnum;
        }  
        [Obsolete("BkgPage is deprecated, please use setBkgPage instead.")]
        [ComVisible(true)]
        public int BkgPage(int pgnum){return setBkgPage(pgnum);}

        [ComVisible(true)]        
        public String setPwdList(String a)
        {          
            setOpt("bkgpass",a);
            return a;
        }      
        [Obsolete("BkgPassword is deprecated, please use setPwdList instead.")]
        [ComVisible(true)]
        public String BkgPassword(String a){return setPwdList(a);}

        [ComVisible(true)]        
        public String setPageBorder(Double width,
            String color = "",
            Double padding = 0)
        {          
            String s = width + "," + color + "," + padding;
            setOpt("border",s);
            return s;
        }      
        [Obsolete("PageBorder is deprecated, please use setPageBorder instead.")]
        [ComVisible(true)]
        public String PageBorder(Double width,
            String color = "",
            Double padding = 0)
        {return setPageBorder(width, color, padding);}

        [ComVisible(true)]
        public String setIniFile(String a = "")
        {
          setOpt("init",a);
          return a;
        }       
        [Obsolete("IniFile is deprecated, please use setIniFile instead.")]
        [ComVisible(true)]
        public String IniFile(String a = ""){return setIniFile(a);}

        [ComVisible(true)]
        public void setBarcodes()
        {
            setOpt("barcodes",true);
        }        
        [Obsolete("Barcodes is deprecated, please use setBarcodes instead.")]
        [ComVisible(true)]
        public void Barcodes(){setBarcodes();}


        [ComVisible(true)]        
        public String setLogFile(String a)
        {          
            setOpt("log",a);
            return a;
        }      

        [ComVisible(true)]        
        public void setNoExtract()
        {          
            setOpt("noextract",true);
        }      
        [Obsolete("NoExtract is deprecated, please use setNoExtract instead.")]
        [ComVisible(true)]
        public void NoExtract() {setNoExtract();}

        [ComVisible(true)]        
        public void setNoFillIn()
        {          
            setOpt("nofillin",true);
        }      
        [Obsolete("NoFillIn is deprecated, please use setNoFillIn instead.")]
        [ComVisible(true)]
        public void NoFillIn() {setNoFillIn();}

        [ComVisible(true)]        
        public void setNoAssemble()
        {          
            setOpt("noassemble",true);
        }      
        [Obsolete("NoAssemble is deprecated, please use setNoAssemble instead.")]
        [ComVisible(true)]
        public void NoAssemble() {setNoAssemble();}

        [ComVisible(true)]        
        public void setNoDigital()
        {          
            setOpt("nodigital",true);
        }     
        [Obsolete("NoDigital is deprecated, please use setNoDigital instead.")]
        [ComVisible(true)]
        public void NoDigital() {setNoDigital();}

        [ComVisible(true)]        
        public void setMargins(Double top = double.MinValue,
            Double right = double.MinValue,
            Double bottom = double.MinValue,
            Double left = double.MinValue)
        {          
            if (top != double.MinValue)
                setOpt("tm",(top * unitsMult).ToString());
            if (left != double.MinValue)
                setOpt("lm",(left * unitsMult).ToString());
            if (bottom != double.MinValue)
                setOpt("bm",(bottom * unitsMult).ToString());
            if (right != double.MinValue)
                setOpt("rm",(right * unitsMult).ToString());
        }      
        [Obsolete("LeftMargin is deprecated, please use setMargins instead.")]
        [ComVisible(true)]
        public Double LeftMargin(Double a){setMargins(left: a);return a;}
        [Obsolete("RightMargin is deprecated, please use setMargins instead.")]
        [ComVisible(true)]
        public Double RightMargin(Double a){setMargins(right: a);return a;}
        [Obsolete("TopMargin is deprecated, please use setMargins instead.")]
        [ComVisible(true)]
        public Double TopMargin(Double a){setMargins(top: a);return a;}
        [Obsolete("BottomMargin is deprecated, please use setMargins instead.")]
        [ComVisible(true)]
        public Double BottomMargin(Double a){setMargins(bottom: a);return a;}

        [ComVisible(true)]        
        public void setColumns(int cols,
            Double spacing = double.MinValue)
        {          
            setOpt("cols",cols + "");
            if (spacing != double.MinValue)
                setOpt("colsp",(spacing * unitsMult).ToString());
        }      
        [Obsolete("Columns is deprecated, please use setColumns instead.")]
        [ComVisible(true)]
        public int Columns(int a){setColumns(cols: a);return a;}
        [Obsolete("ColumnSpace is deprecated, please use setColumns instead.")]
        [ComVisible(true)]
        public Double ColumnSpace(Double a){setOpt("colsp",(a * unitsMult).ToString());return a;}

        [ComVisible(true)]
        public Double setLineSpace(Double x)
        {
            setOpt("lnspace",x.ToString());
          return x;
        }       
        [Obsolete("LineSpace is deprecated, please use setLineSpace instead.")]
        [ComVisible(true)]
        public Double LineSpace(Double x){return setLineSpace(x);}

        [ComVisible(true)]
        public String setPageTxtFont(String font,
            Double pointSize)
        {
            String s = font + "," + pointSize;
            setOpt("pagetxtfont",s);
          return s;
        }       
        [Obsolete("PageTxtFont is deprecated, please use setPageTxtFont instead.")]
        [ComVisible(true)]
        public String PageTxtFont(String font,
            Double pointSize){return setPageTxtFont(font, pointSize);}

        [ComVisible(true)]
        public String setPageNumStr(String a,
            int pg = 0)
        {
          setOpt("pagestr" + (pg > 0 ? pg + "" : ""),a);
          return a;
        }       
        [Obsolete("PageStr is deprecated, please use setPageNumStr instead.")]
        [ComVisible(true)]
        public String PageStr(String a,
            int pg = 0){return setPageNumStr(a, pg);}

        [ComVisible(true)]
        public String setPageNumPos(Double x,
            Double y,
            int pg = 0)
        {
          setOpt("pagepos" + (pg > 0 ? pg + "" : ""),(x * unitsMult) + "," + (y * unitsMult));
          return x + "," + y;
        }       
        [Obsolete("PagePos is deprecated, please use setPageNumPos instead.")]
        [ComVisible(true)]
        public String PagePos(Double x,
            Double y,
            int pg = 0){return setPageNumPos(x, y, pg);}

        [ComVisible(true)]
        public String setPageNumFont(String fontName,
            Double ptSize,
            int pg = 0)
        {
          setOpt("pagenumfont" + (pg > 0 ? pg + "" : ""),ptSize == 0 ? fontName : fontName + "," + ptSize);
          return fontName;
        }       
        [Obsolete("PageNumFont is deprecated, please use setPageNumFont instead.")]
        [ComVisible(true)]
        public String PageNumFont(String fontName,
            Double ptSize,
            int pg = 0){return setPageNumFont(fontName, ptSize, pg);}

        [ComVisible(true)]
        public void setPageNum(String format,
            String fontName,
            Double ptSize,
            Double x,
            Double y,
            int pg = 0)
        {
            setPageNumStr(format, pg);
            setPageNumFont(fontName, ptSize, pg);
            setPageNumPos(x, y, pg);
        }       

        [ComVisible(true)]
        public String setPre(String a = "")
        {
            a = a.ToLower();
            switch (a) {
                case "plain": case "html": case "white": case "whitenb":
                    setOpt("pre",a);
                    break;
                default:
                    setOpt("pre",true);
                    break;
                }                
          return a;
        }       
        [Obsolete("Pre is deprecated, please use setPre instead.")]
        [ComVisible(true)]
        public String Pre(String a = ""){return setPre(a);}

        [ComVisible(true)]
        public String setUTF8(String lang, 
            String encode = "")
        {
            String s = lang + (encode.Equals("") ? "" : "," + encode);
            setOpt("pre",s);
            return s;
        }       
        [Obsolete("UTF8 is deprecated, please use setUTF8 instead.")]
        [ComVisible(true)]
        public String UTF8(String lang, String encode = ""){return setUTF8(lang, encode);}

        [ComVisible(true)]
        public String setStrokeColor(String a = "")
        {
          setOpt("scolor",a);
          return a;
        }       
        [Obsolete("SColor is deprecated, please use setStrokeColor instead.")]
        [ComVisible(true)]
        public String SColor(String a = ""){return setStrokeColor(a);}

        [ComVisible(true)]
        public String setFillColor(String a)
        {
          setOpt("scolor",a);
          return a;
        }       
        [Obsolete("FColor is deprecated, please use setFillColor instead.")]
        [ComVisible(true)]
        public String FColor(String a){return setFillColor(a);}

        [ComVisible(true)]
        public void setColOne()
        {
            setOpt("c1",true);
        }
        [Obsolete("ColOne is deprecated, please use setColOne instead.")]
        [ComVisible(true)]
        public void ColOne(){setColOne();}

        [ComVisible(true)]
        public void setNoSoftBreak()
        {
            setOpt("nosoftbreak",true);
        }
        [Obsolete("NoSoftBreak is deprecated, please use setNoSoftBreak instead.")]
        [ComVisible(true)]
        public void NoSoftBreak(){setNoSoftBreak();}

        [ComVisible(true)]
        public void setRemoveCtl()
        {
            setOpt("removectl",true);
        }
        [Obsolete("RemoveCtl is deprecated, please use setRemoveCtl instead.")]
        [ComVisible(true)]
        public void RemoveCtl(){setRemoveCtl();}

        [ComVisible(true)]
        public void setPrintDlg()
        {
            setOpt("printdlg",true);
        }
        [Obsolete("UsePrintDlg is deprecated, please use setPrintDlg instead.")]
        [ComVisible(true)]
        public void UsePrintDlg(){setPrintDlg();}

        [ComVisible(true)]
        public String setPrinter(String printer,
            String device = "",
            String port = "")
        {
            String s = printer + (device.Equals("") ? "" : "," + device) + (port.Equals("") ? "" : "," + port);
            setOpt("printer",s);
            return s;
        }
        [Obsolete("UsePrinter is deprecated, please use setPrinter instead.")]
        [ComVisible(true)]
        public String UsePrinter(String printer,
            String device = "",
            String port = ""){return setPrinter(printer, device, port);}

        [ComVisible(true)]
        public int setNumCopies(int a)
        {
            if (a > 0){
                setOpt("copies",a + "");
            }
            return a;
        }
        [Obsolete("NumCopies is deprecated, please use setNumCopies instead.")]
        [ComVisible(true)]
        public int NumCopies(int a){return setNumCopies(a);}

        [ComVisible(true)]
        public int setRenderMode(int a)
        {
            setOpt("rend",a + "");
            return a;
        }
        [Obsolete("Render is deprecated, please use setRenderMode instead.")]
        [ComVisible(true)]
        public int Render(int a){return setRenderMode(a);}

        [ComVisible(true)]
        public int setTabSpace(int a)
        {
            setOpt("tabspace",a + "");
            return a;
        }
        [Obsolete("TabSpace is deprecated, please use setTabSpace instead.")]
        [ComVisible(true)]
        public int TabSpace(int a){return setTabSpace(a);}

        [ComVisible(true)]
        public int setTabAbsSpace(int a)
        {
            setOpt("tababsspace",a + "");
            return a;
        }
        [Obsolete("TabAbsSpace is deprecated, please use setTabAbsSpace instead.")]
        [ComVisible(true)]
        public int TabAbsSpace(int a){return setTabAbsSpace(a);}

        [ComVisible(true)]
        public char setAlign(char a)
        {
            if (char.ToString(char.ToLower(a)).IndexOfAny(new char[] {'l','r','c','j'}) != -1){                
                setOpt("align",a);
            }
          return a;
        }       
        [Obsolete("Align is deprecated, please use setAlign instead.")]
        [ComVisible(true)]
        public char Align(char a){return setAlign(a);}

        // Calls buildPDF or buildPDFTCP
        [ComVisible(true)]
        public object buildPDF(bool waitForExit = true,
            String saveFile = "")    
        {   
            object host;
            if (server.TryGetValue("host", out host)
                || servers.Count > 0) {
                // if there is a server or servers, build using TCP                 
                // waitForExit means return the byte array of the PDF
                return buildPDFTCP(waitForExit, saveFile);
            } else {
                // otherwise, build using the executable
                if (!saveFile.Equals("")){
                    setOutFile(saveFile); // shorthand for calling setOutFile
                }
                return build(waitForExit);
            }
        }

        // Call the executable (non server mode)
        private object build(bool waitForExit = true)
        {    
            byte[] bytes = {};
            String errMsg = "";
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.CreateNoWindow = false;
            startInfo.UseShellExecute = !waitForExit;
            if (waitForExit){
                startInfo.RedirectStandardOutput = true;
                startInfo.RedirectStandardInput = true;
            }
            startInfo.FileName = exe;
            startInfo.WindowStyle = ProcessWindowStyle.Hidden;
            startInfo.Arguments = setBaseOpts();  
            (bytes, errMsg) = runProcess(startInfo,waitForExit);
            Results ret = new Results();
            ret.Bytes = bytes;
            ret.Msg = errMsg.Equals("") ? "OK" : errMsg;
            return ret;
        }

        // Reset the options
        [ComVisible(true)]
        public void resetOpts(bool resetServer = false){
            cmds = new List<string>();
            opts.Clear();
            unitsMult = 1;
            if (resetServer){
                server.Clear();
            }
            
        }

        // Passes file to server over socket
        public String sendFileTCP(String fileName,
            String filePath = "",
            byte[] bytes = null){
            String errMsg = "";
            Byte[] data;
            Object host = "";
            // String to store the response ASCII representation.
            String responseData = String.Empty;
            String message = "";
            if (!isServerRunning()){
                return "Server not running";
            }
            try {                         
                // Send a file
                byte[] buffer = new byte[1024];
                int bytesRead = 0;

                message = " -send --binaryname--" + fileName + "--binarybegin--";
                data = System.Text.Encoding.UTF8.GetBytes(message);   
                stream.Write(data, 0, data.Length); 
                if (bytes != null && bytes.Length > 0){
                    stream.Write(bytes, 0, bytes.Length);                    
                } else if (filePath.Equals("--input--commands--")){ 
                    foreach (String value in cmds){
                        data = System.Text.Encoding.UTF8.GetBytes(value);
                        stream.Write(data, 0, data.Length);
                    }
                }   
                else            
                if (!filePath.Equals("")){
                    BinaryReader br;
                    br = new BinaryReader(new FileStream(filePath, FileMode.Open));

                    while ((bytesRead = br.Read(buffer, 0, buffer.Length)) > 0)
                        stream.Write(buffer, 0, bytesRead);                    

                }
                message = "--binaryend-- ";
                data = System.Text.Encoding.UTF8.GetBytes(message);             
                stream.Write(data, 0, data.Length);                           

            } catch (SocketException e) {
                errMsg = e.Message;
            } catch (IOException e) {
                errMsg = e.Message;
            }
            return (errMsg);
        }

        private void setOpt(String k, object v){
            opts[k] = v;
        }

        private (byte[], String) runProcess(ProcessStartInfo startInfo, bool waitForExit) {
            // Start the process with the info we specified.
            // Call WaitForExit if we are waiting for process to complete. 
            byte[] bytes = {};
            try {              
                using (Process exeProcess = Process.Start(startInfo))
                {
                    if (waitForExit){
                        StreamWriter inputCmds = exeProcess.StandardInput;
                        foreach (String value in cmds){
                            inputCmds.Write(value);
                        }
                        inputCmds.Close();
                    }

                    MemoryStream memstream = new MemoryStream();
                    byte[] buffer = new byte[1024];
                    int bytesRead = 0;
                    if (waitForExit){
                        BinaryReader br = new BinaryReader(exeProcess.StandardOutput.BaseStream);
                        while ((bytesRead = br.Read(buffer, 0, buffer.Length)) > 0)
                            memstream.Write(buffer, 0, bytesRead);                    
                    }                        
                    if (waitForExit){
                        exeProcess.WaitForExit();
                        bytes = memstream.ToArray();
                    }
                }
            }
            catch (Exception e){
                return (bytes, e.Message);
            }
            return (bytes, "");

        }

        // Passes data to server over socket but does not finalize 
        // (that is, does not send BUILDPDF command)
        private String sendTCP(){
            String errMsg = "";
            Byte[] data;
            Object host = "";
            // String to store the response ASCII representation.
            String responseData = String.Empty;
            object message = "";

            try {         
                
                object s;
                if (opts.TryGetValue("serverCmd", out s)){
                    message = s;
                    opts.Remove("serverCmd");
                } else {
                    message = setBaseOpts();
                }
                // Send commands
                data = System.Text.Encoding.UTF8.GetBytes((String) message);             
                stream.Write(data, 0, data.Length);      

            } catch (SocketException e) {
                errMsg = e.Message;
            } catch (IOException e) {
                errMsg = e.Message;
            }
            return (errMsg);
        }

        // build the command line string to pass to the executable
        private String setBaseOpts(){
            object s = "";
            String message = "";            
            if (opts.TryGetValue("inFile", out s))
                message += " \"" + s + "\"";
            else if (cmds.Count > 0)        
                    message += " -"; // get from stdin
            if (opts.TryGetValue("outFile", out s)) {
                if (!s.Equals(""))
                    message += " \"" + s + "\"";
            }
            foreach (KeyValuePair<string,object> opt in opts){
                if (!opt.Key.Equals("inFile") && !opt.Key.Equals("outFile")){
                    if (opt.Key.Equals("extOpts")){
                        message += " " + opt.Value + " ";
                    } else {
                        message += " -" + opt.Key;
                        if (!(opt.Value is bool)){
                            message += " \"" + opt.Value.ToString().Replace("\"","\\\"") + "\"";
                        }
                    }                
                }
            }               
            return message;                         
        }
    
        // Send the BUILDPDF command to server to run the commands
        private (byte[], String, int) callTCP(bool isStop = false,
            bool isStatus = false,
            bool retBytes = false,
            String saveFile = ""){
            String errMsg = "";
            Byte[] data = {};
            Byte[] bytes = {};
            int numPages = 0;
            Object host = "";
            bool retPDF = retBytes;
            // String to store the response ASCII representation.
            String responseData = String.Empty;

            if (!isServerRunning()){
                return (bytes, "Server not running", 0);
            }

            errMsg = sendTCP();
            if (errMsg.Equals("")){
                if (opts.TryGetValue("autosend", out object s))
                  retPDF = true; // need to keep open and send files
            }
            if (!saveFile.Equals("")){
                retPDF = true; // need to save the PDF
            }

            try {         
                if (retPDF){
                    data = System.Text.Encoding.ASCII.GetBytes(" -return ");             
                    stream.Write(data, 0, data.Length);      
                }
                data = System.Text.Encoding.ASCII.GetBytes("\nBUILDPDF\n");             
                stream.Write(data, 0, data.Length);      
                if (isStatus){ 
                    do {
                        data = new Byte[1024];
                        // Read the first batch of the TcpServer response bytes.
                        Int32 rawData = stream.Read(data, 0, data.Length);
                        responseData += System.Text.Encoding.ASCII.GetString(data, 0, rawData);
                    }
                    while (stream.DataAvailable);
                } else if (!isStop) {
                    MemoryStream memstream = new MemoryStream();
                    Socket s = client.Client;
                    byte[] buffer = new byte[1024];                    
                    int bytesRead = 0;
                    String retStr = "";
                    String[] retCmds = new String[2];                    
                    
                    while(true) {
                        // Read the first batch of the TcpServer response bytes.
                        buffer = new byte[1024];
                        bytesRead = stream.Read(buffer, 0, buffer.Length);
                        retStr = System.Text.Encoding.ASCII.GetString(buffer, 0, buffer.Length);
                        if (retStr.ToLower().StartsWith("content-length:")){
                          // Receiving the PDF back
                          while ((bytesRead = stream.Read(buffer, 0, buffer.Length)) > 0){
                              memstream.Write(buffer, 0, bytesRead);     
                          }                        
                          bytes = memstream.ToArray();
                          memstream.SetLength(0);
                          break;
                        } else if (retStr.ToLower().StartsWith("page-count:")){
                          retCmds = retStr.Split(new string[] { "\n\n", "\n" }, StringSplitOptions.None);
                          retCmds = retCmds[0].Split(new string[] { ":" }, (Int32) 2, StringSplitOptions.None);
                          try {
                            numPages = Int32.Parse(retCmds[1]);
                          } catch (FormatException) {
                            numPages = -1;
                          }
                        } else {
                          retCmds = retStr.Split(new string[] { "\n\n", "\n" }, StringSplitOptions.None);
                          retCmds = retCmds[0].Split(new string[] { ":" }, (Int32) 2, StringSplitOptions.None);                           
                          if (retCmds.Length > 0 && 
                              (retCmds[0].ToLower().Equals("send-file")
                              || retCmds[0].ToLower().Equals("send-md5"))) {

                            if (retCmds[0].ToLower().Equals("send-file")){
                                retCmds[1] = retCmds[1].Trim();
                                FileInfo f = new FileInfo(retCmds[1]);
                                data = System.Text.Encoding.ASCII.GetBytes("Content-Length: " + f.Length + "\n\n");
                                stream.Write(data, 0, data.Length);
                                BinaryReader br = new BinaryReader(new FileStream(retCmds[1], FileMode.Open));

                                while ((bytesRead = br.Read(buffer, 0, buffer.Length)) > 0)
                                    stream.Write(buffer, 0, bytesRead);     
                                stream.Flush();
                            } else {
                                    String message = "";
                                    MD5 md5 = MD5.Create();
                                    FileStream fStream = File.OpenRead(retCmds[1]);
                                    byte[] md5Bytes = md5.ComputeHash(fStream);
                                    String fileHash = ByteArrayToString(md5Bytes);
                                    data = System.Text.Encoding.ASCII.GetBytes(message);             
                                    stream.Write(data, 0, data.Length);
                                    stream.Flush();
                                }
                            }
                            if ((s.Poll(1000, SelectMode.SelectRead) && (s.Available == 0)) || !s.Connected)
                            {
                                break;
                            } 
                        }                     
                    }
                }                
                if (!saveFile.Equals("")){
                    File.WriteAllBytes (saveFile, bytes);
                    Array.Clear(bytes, 0, bytes.Length);
                }
                if (!retBytes){
                    Array.Copy(bytes, new byte[0], 0);
                }
            } catch (SocketException e) {
                errMsg = e.Message;
            } catch (IOException e) {
                errMsg = e.Message;                
            }
            stream.Close();
            client.Close();

            return (bytes, (responseData.Equals("") ? errMsg : responseData), numPages);
        }

        private static string ByteArrayToString(byte[] ba)
        {
            return BitConverter.ToString(ba).Replace("-","");
        }

        private static byte[] ConvertHexToByteArray(string hexString)
        {
            byte[] byteArray = new byte[hexString.Length / 2];
    
            for (int index = 0; index < byteArray.Length; index++)
            {
                string byteValue = hexString.Substring(index * 2, 2);
                byteArray[index] = byte.Parse(byteValue, NumberStyles.HexNumber, CultureInfo.InvariantCulture);
            }
    
            return byteArray;
        }
         
    }
}

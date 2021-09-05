
using System;
 using System.Collections.Generic;
 using System.Linq;
 using System.Text;
 using System.Net.Sockets;
 using System.IO;
using System.Windows.Forms;

namespace ConnTelnetClient
{
    class Program
    {
       
       public static string ip;
         
       public static FileStream fs;
        public static StreamWriter sw;

       public static  Queue<string> ipQueue = new Queue<string>();
        public static List<string> commandList = new List<string>();
         static void Main(string[] args)
         {
            FileStream commandFile = new FileStream(System.Windows.Forms.Application.StartupPath + "\\command.ini", FileMode.Open);
            StreamReader commandSr = new StreamReader(commandFile);
           // commandList.Add("mo");
            //commandList.Add("motorola");
            commandList.Add("fa");
            commandList.Add("factory");
            string line;
            while ((line = commandSr.ReadLine()) != null)
            {
                commandList.Add(line);

            }
            commandSr.Close();
            commandFile.Close();



            FileStream ipFile = new FileStream(System.Windows.Forms.Application.StartupPath + "\\ip.ini", FileMode.Open);
            StreamReader ipSr = new StreamReader(ipFile);
            while ((line = ipSr.ReadLine())!=null)
            {
                ipQueue.Enqueue(line);
            }
            ipSr.Close();
            ipFile.Close();

            DateTime dt = DateTime.Now;
            string name = dt.Year + "年" + dt.Month + "月" + dt.Day + "日" + dt.Hour + "时" + dt.Minute + "分.txt";
            fs = new FileStream(System.Windows.Forms.Application.StartupPath+"\\"+name, FileMode.OpenOrCreate);
            sw = new StreamWriter(fs);

            sw.WriteLine("start up");

            while (ipQueue.Count>0)
            {
                Run2(ipQueue.Dequeue());
            }
            sw.Flush();
            sw.Close();
            fs.Close();

            Console.WriteLine("结束,Close after 60 second");
            System.Threading.Thread.Sleep(60000);
           
        }

      

        static public void Run2(string ip)
        {
         NetworkStream stream;
         TcpClient tcpclient;
        
            try
            {
                
                Console.WriteLine(ip + " 连接...");
                tcpclient = new TcpClient(ip, 23);  // 连接服务器
                stream = tcpclient.GetStream();   // 获取网络数据流对象
                SendKeys.SendWait("{Enter}");
            }
            catch (Exception e)
            {
                
                sw.WriteLine(ip + " 无法连接");
                Console.WriteLine(ip + " 无法连接");
                return;
            }
           
            StreamWriter sw_net = new StreamWriter(stream);
            StreamReader sr_net = new StreamReader(stream);
            sw.WriteLine(ip + " start:");
            int index = 0;
            while (true)
            {
                //Read Echo
                //Set ReadEcho Timeout
                stream.ReadTimeout = 10;
                try
                {
                    while (true)
                    {
                        char c = (char)sr_net.Read();
                        if (c < 1024)
                        {
                            if (c == 27)
                            {
                                while (sr_net.Read() != 109) { }
                            }
                            else
                            {
                                Console.Write(c);
                                sw.Write(c);
                            }
                        }
                    }
                }
                catch (Exception e)
                {
                    sw.WriteLine("========================================");
                    //Console.WriteLine(e.Message);
                }


                try
                {
                    //Send CMD

                    if (index<commandList.Count)
                    {
                        sw_net.Write("{0}\r\n", commandList[index]);
                        SendKeys.SendWait("{Enter}");
                        sw_net.Flush();
                        //SendKeys.SendWait("{Enter}");

                    }
                    
                    //if (index == 1)
                    //{
                    //    sw_net.Write("{0}\r\n", "mo");
                    //    SendKeys.SendWait("{Enter}");
                    //    sw_net.Flush();
                    //    SendKeys.SendWait("{Enter}");

                    //}
                    //else if (index == 2)
                    //{
                    //    sw_net.Write("{0}\r\n", "motorola");
                    //    SendKeys.SendWait("{Enter}");
                    //    sw_net.Flush();
                    //    SendKeys.SendWait("{Enter}");
                    //}
                    //else if (index == 3)
                    //{
                    //    sw_net.Write("{0}\r\n", "dpm 1 get vswr");
                    //    SendKeys.SendWait("{Enter}");
                    //    sw_net.Flush();
                    //    SendKeys.SendWait("{Enter}");
                    //}
                    //else if (index == 4)
                    //{
                    //    sw_net.Write("{0}\r\n", "exit");
                    //    SendKeys.SendWait("{Enter}");
                    //    sw_net.Flush();
                    //    SendKeys.SendWait("{Enter}");
                    //}
                    System.Threading.Thread.Sleep(300);
                    index++;
                    if (index==10)
                    {
                        sw.WriteLine("=====================================End===============================================");
                        sw_net.Close();
                        sr_net.Close();
                        //stream.Close();
                        tcpclient.Close();
                        break;
                    }
                }
                catch (Exception e)
                {

                    Console.WriteLine(e.Message);
                }
                
              
            }

        }
    }
}

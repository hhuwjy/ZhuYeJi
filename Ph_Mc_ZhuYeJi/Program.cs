using HslCommunication;
using HslCommunication.LogNet;
using HslCommunication.Profinet.Omron;
using System.Net.Sockets;
using System.Reflection;
using System.Threading;
using System.Timers;
using Timer = System.Timers.Timer;
using System.Net;
using Newtonsoft.Json.Linq;
using HslCommunication.MQTT;
using HslCommunication.Profinet.Siemens;
using HslCommunication.Profinet.OpenProtocol;
using NPOI.XSSF.UserModel;
using static Ph_Mc_ZhuYeJi.UserStruct;
using Arp.Plc.Gds.Services.Grpc;
using Opc.Ua;
using Microsoft.Extensions.Logging;
using Google.Protobuf.WellKnownTypes;
using NPOI.POIFS.Crypt.Dsig;
using Grpc.Core;
using static Arp.Plc.Gds.Services.Grpc.IDataAccessService;
using Grpc.Net.Client;
using static Ph_Mc_ZhuYeJi.GrpcTool;
using static Ph_Mc_ZhuYeJi.KeyenceComm;
using System.Text;
using HslCommunication.CNC.Fanuc;
using System.Security.Claims;
using NPOI.Util;
using System.Text.RegularExpressions;
using SixLabors.ImageSharp;
using NPOI.SS.Formula.Functions;
using HslCommunication.Instrument.DLT;
using SixLabors.ImageSharp.Processing;
using System.Drawing;
using System.Net.NetworkInformation;
using Ph_Mc_ZhuYeJi;
using HslCommunication.Profinet.Keyence;



namespace Ph_Mc_ZhuYeJi
{
    class Program
    {
        /// <summary>
        /// app初始化
        /// </summary>
        /// 

        // 创建日志
        const string logsFile = ("/opt/plcnext/apps/ZhuYeAppLogs");
        //const string logsFile = "D:\\2024\\Work\\12-冠宇数采项目\\ReadFromStructArray\\ZhuYeJi_MC\\ZhuYeJi_MC_Log";

        public static ILogNet logNet = new LogNetFileSize(logsFile, 5 * 1024 * 1024); //限制了日志大小

        //创建Grpc实例
        public static GrpcTool grpcToolInstance = new GrpcTool();

        //设置grpc通讯参数
        public static CallOptions options1 = new CallOptions(
                new Metadata {
                        new Metadata.Entry("host","SladeHost")
                },
                DateTime.MaxValue,
                new CancellationTokenSource().Token);
        public static IDataAccessServiceClient grpcDataAccessServiceClient = null;

        //创建Tool的API实例
        public static ToolAPI tool = new ToolAPI();

        //创建nodeID字典 (读取XML用）
        public static Dictionary<string, string> nodeidDictionary1;
        public static Dictionary<string, string> nodeidDictionary2;

        //读取Excel用
        static ReadExcel readExcel = new ReadExcel();

        //MC Client实例化 
        public static KeyenceComm keyenceClients = new KeyenceComm();
        static int clientNum = 2;  //一个EPC对应采集两个基恩士的数据（点表相同）  上位链路+MC协议，同时在线加起来不能超过15台
        public static KeyenceMcNet[] _mc = new KeyenceMcNet[clientNum];
        
        
        //PLC IP Address地址
        public static List<string> plcIpAddresses = new List<string>();

        //开启线程
        static int thrNum = 7;  
        static Thread[] thr = new Thread[thrNum];

        #region 数据点位名和设备总览表格的实例化结构体


        // 设备总览
        static DeviceInfoStruct_IEC[] deviceInfoStruct1_IEC;
        static DeviceInfoStruct_IEC[] deviceInfoStruct2_IEC;
        #endregion


        #region 从Excel解析来的数据实例化  (10MEJQ33-4793   (10MEJQ33-4790))

        // 设备信息（47个工位）

        static DeviceInfoConSturct_MC[] Battery_Memory;
        static DeviceInfoConSturct_MC[] Battery_Clear;
        static DeviceInfoConSturct_MC[] Device_Enable;
        static DeviceInfoDisSturct_MC[] EarCode;
        static DeviceInfoDisSturct_MC[] BarCode; 


        //1000ms数据
        static OneSecInfoStruct_MC[] Function_Enable;
        static OneSecInfoStruct_MC[] Production_Data;
        static OneSecInfoStruct_MC[] OEE1;
        static OneSecInfoStruct_MC[] OEE2;

        static OneSecInfoStruct_MC[] Life_Management;

        static OneSecInfoStruct_MC[] Alarm1;
        static OneSecInfoStruct_MC[] Alarm2;
        static OneSecInfoStruct_MC[] Alarm3;
        static OneSecInfoStruct_MC[] Alarm4;


        //三大工位

        static StationInfoStruct_MC[] ZhuYeWei;
        static StationInfoStruct_MC[] JingZhiWei;
        static StationInfoStruct_MC[] FengZhuangWei;


        #endregion

  

        static void Main(string[] args)
        {
            int stepNumber = 5;


            List<WriteItem> listWriteItem = new List<WriteItem>();
            IDataAccessServiceReadSingleRequest dataAccessServiceReadSingleRequest = new IDataAccessServiceReadSingleRequest();

            bool isThreadZeroRunning = false;
            bool isThreadOneRunning = false;
            bool isThreadTwoRunning = false;
            bool isThreadThreeRunning = false;
            bool isThreadFourRunning = false;
            bool isThreadFiveRunning = false;


            int IecTriggersNumber = 0;

            //采集值缓存区，需要写入Excel
            AllDataReadfromMC allDataReadfromMC_4793 = new AllDataReadfromMC();
            AllDataReadfromMC allDataReadfromMC_4790 = new AllDataReadfromMC();




            // 4793 为设备一   4790 为设备二
            while (true)
            {
                switch (stepNumber)
                {
                    case 5:
                        {
                            #region Grpc连接

                            var udsEndPoint = new UnixDomainSocketEndPoint("/run/plcnext/grpc.sock");
                            var connectionFactory = new UnixDomainSocketConnectionFactory(udsEndPoint);

                            //grpcDataAccessServiceClient
                            var socketsHttpHandler = new SocketsHttpHandler
                            {
                                ConnectCallback = connectionFactory.ConnectAsync
                            };
                            try
                            {
                                GrpcChannel channel = GrpcChannel.ForAddress("http://localhost", new GrpcChannelOptions  // Create a gRPC channel to the PLCnext unix socket
                                {
                                    HttpHandler = socketsHttpHandler
                                });
                                grpcDataAccessServiceClient = new IDataAccessService.IDataAccessServiceClient(channel);// Create a gRPC client for the Data Access Service on that channel
                            }
                            catch (Exception e)
                            {
                                Console.WriteLine("ERRO: {0}", e);
                                //logNet.WriteError("Grpc connect failed!");
                            }
                            #endregion

                            stepNumber = 6;

                        }
                        break;


                    case 6:
                        {

                            #region 从xml获取nodeid，Grpc发送到对应变量时使用，注意xml中的别名要和对应类的属性名一致 
                            try
                            {
                                const string filePath1 = "/opt/plcnext/apps/GrpcSubscribeNodes_4793.xml";             //EPC中存放的路径  
                                //const string filePath1 = "D:\\2024\\Work\\12-冠宇数采项目\\ReadFromStructArray\\ZhuYeJi_MC\\ZhuYeJi\\Ph_Mc_ZhuYeJi\\GrpcSubscribeNodes\\GrpcSubscribeNodes_4793.xml";  //PC中存放的路径 

                                nodeidDictionary1 = grpcToolInstance.getNodeIdDictionary(filePath1);  //将xml中的值写入字典中
                                logNet.WriteInfo("NodeID Sheet 4793 文件读取成功");

                            }
                            catch (Exception e)
                            {                              
                                logNet.WriteError("NodeID Sheet 4793 文件读取失败");
                                logNet.WriteError("Error:" + e);
                            }

                            try
                            {
                                const string filePath2 = "/opt/plcnext/apps/GrpcSubscribeNodes_4790.xml";             //EPC中存放的路径  
                                //const string filePath2 = "D:\\2024\\Work\\12-冠宇数采项目\\ReadFromStructArray\\ZhuYeJi_MC\\ZhuYeJi\\Ph_Mc_ZhuYeJi\\GrpcSubscribeNodes\\GrpcSubscribeNodes_4790.xml";  //PC中存放的路径 

                                nodeidDictionary2 = grpcToolInstance.getNodeIdDictionary(filePath2);  //将xml中的值写入字典中
                                logNet.WriteInfo("NodeID Sheet 4790 文件读取成功");

                            }
                            catch (Exception e)
                            {
                                logNet.WriteError("NodeID Sheet 4790 文件读取失败");
                                logNet.WriteError("Error:" + e);
                            }


                            #endregion

                            stepNumber = 10;
                        }


                        break;


                    case 10:
                    { 
                            /// <summary>
                            /// 执行初始化
                            /// </summary>


                            #region 读取Excel 

                            logNet.WriteError("注液机设备数采APP已启动");

                            // 读取 设备4793 的Excel

                            //string excelFilePath1 = Directory.GetCurrentDirectory() + "\\ZYJData(4793).xlsx";     //PC端测试路径
                            string excelFilePath1 = "/opt/plcnext/apps/ZYJData(4793).xlsx";                         //EPC存放路径

                            XSSFWorkbook excelWorkbook1 = readExcel.connectExcel(excelFilePath1);
                            Console.WriteLine("ZYJData(4793) 读取 {0}", excelWorkbook1 != null ? "成功" : "失败");
                            logNet.WriteInfo("ZYJData(4793) 读取 ", excelWorkbook1 != null ? "成功" : "失败");


                            // 读取 设备4790 的Excel

                            //string excelFilePath2 = Directory.GetCurrentDirectory() + "\\ZYJData(4790).xlsx";     //PC端测试路径
                            string excelFilePath2 = "/opt/plcnext/apps/ZYJData(4790).xlsx";                         //EPC存放路径

                            XSSFWorkbook excelWorkbook2 = readExcel.connectExcel(excelFilePath2);
                            Console.WriteLine("ZYJData(4790) 读取 {0}", excelWorkbook2 != null ? "成功" : "失败");
                            logNet.WriteInfo("ZYJData(4790) 读取 ", excelWorkbook2 != null ? "成功" : "失败");



                            // 给IEC发送 Excel读取成功的信号
                            var tempFlag_finishReadExcelFile = true;

                            listWriteItem.Clear();
                            listWriteItem.Add(grpcToolInstance.CreatWriteItem(nodeidDictionary1["flag_finishReadExcelFile"], Arp.Type.Grpc.CoreType.CtBoolean, tempFlag_finishReadExcelFile));
                            if (grpcToolInstance.WriteDataToDataAccessService(grpcDataAccessServiceClient, grpcToolInstance.ServiceWriteRequestAddDatas(listWriteItem.ToArray()), new IDataAccessServiceWriteResponse(), options1))
                            {
                                //Console.WriteLine("{0}      flag_finishReadExcelFile写入IEC: success", DateTime.Now);
                                logNet.WriteInfo("[Grpc]", "flag_finishReadExcelFile 写入IEC成功");
                            }
                            else
                            {
                                //Console.WriteLine("{0}      flag_finishReadExcelFile写入IEC: fail", DateTime.Now);
                                logNet.WriteError("[Grpc]", "flag_finishReadExcelFile 写入IEC失败");
                            }


                            #endregion


                            #region 将readExcel变量中的值，写入对应的实例化结构体中（4793  4790） 

                            //设备信息
                            Battery_Memory = readExcel.ReadOneDeviceInfoConSturctInfo_Excel(excelWorkbook1, "设备信息", "电芯记忆信号（BOOL)");
                            Battery_Clear = readExcel.ReadOneDeviceInfoConSturctInfo_Excel(excelWorkbook1, "设备信息", "电芯记忆清除按钮（BOOL)");
                            Device_Enable = readExcel.ReadOneDeviceInfoConSturctInfo_Excel(excelWorkbook1, "设备信息", "工位加工中信号（BOOL）");
                            BarCode = readExcel.ReadOneDeviceInfoDisSturctInfo_Excel(excelWorkbook1, "设备信息", "电芯条码地址（DINT)");
                            EarCode = readExcel.ReadOneDeviceInfoDisSturctInfo_Excel(excelWorkbook1, "设备信息", "极耳码地址(DINT）");

                             //1000ms数据
                            Function_Enable = readExcel.ReadOneSecInfo_Excel(excelWorkbook1, "功能开关",false);
                            Production_Data = readExcel.ReadOneSecInfo_Excel(excelWorkbook1, "生产统计", false);
                            Life_Management = readExcel.ReadOneSecInfo_Excel(excelWorkbook1, "寿命管理", false);

                            OEE1 = readExcel.ReadOneSecInfo_Excel(excelWorkbook1, "OEE(1)", false);
                            OEE2 = readExcel.ReadOneSecInfo_Excel(excelWorkbook1, "OEE(2)", false);
                            Alarm1 = readExcel.ReadOneSecInfo_Excel(excelWorkbook1, "报警信号1", false);
                            Alarm2 = readExcel.ReadOneSecInfo_Excel(excelWorkbook1, "报警信号2", false);
                            Alarm3 = readExcel.ReadOneSecInfo_Excel(excelWorkbook1, "报警信号3", false);
                            Alarm4 = readExcel.ReadOneSecInfo_Excel(excelWorkbook1, "报警信号4",true);

                            //三大工位
                            ZhuYeWei = readExcel.ReadStationInfo_Excel(excelWorkbook1, "加工工位（注液位）");
                            JingZhiWei = readExcel.ReadStationInfo_Excel(excelWorkbook1, "加工工位（静置位）");
                            FengZhuangWei = readExcel.ReadStationInfo_Excel(excelWorkbook1, "加工工位（封装位）");

                            #endregion


                            #region 读取并发送Excel里的设备总览表

                            // 单独设备总览表 + 发送
                            deviceInfoStruct1_IEC = readExcel.ReadDeviceInfo_Excel(excelWorkbook1, "注液机设备总览");  //4793
                            deviceInfoStruct2_IEC = readExcel.ReadDeviceInfo_Excel(excelWorkbook2, "注液机设备总览");  //4790

                            listWriteItem = new List<WriteItem>();

                            //4793
                            try
                            {
                                listWriteItem.Add(grpcToolInstance.CreatWriteItem(nodeidDictionary1["OverviewInfo"], Arp.Type.Grpc.CoreType.CtStruct, deviceInfoStruct1_IEC[0]));
                                var writeItemsArray = listWriteItem.ToArray();
                                var dataAccessServiceWriteRequest = grpcToolInstance.ServiceWriteRequestAddDatas(writeItemsArray);
                                bool result = grpcToolInstance.WriteDataToDataAccessService(grpcDataAccessServiceClient, dataAccessServiceWriteRequest, new IDataAccessServiceWriteResponse(), options1);
                            }
                            catch (Exception e)
                            {
                                Console.WriteLine("ERRO: {0}", e);
                                logNet.WriteError("设备编号4793的设备总览信息发送失败，错误原因 : " + e.ToString());
                            }
                            listWriteItem.Clear();


                            //4790
                            try
                            {
                                listWriteItem.Add(grpcToolInstance.CreatWriteItem(nodeidDictionary2["OverviewInfo"], Arp.Type.Grpc.CoreType.CtStruct, deviceInfoStruct2_IEC[0]));
                                var writeItemsArray = listWriteItem.ToArray();
                                var dataAccessServiceWriteRequest = grpcToolInstance.ServiceWriteRequestAddDatas(writeItemsArray);
                                bool result = grpcToolInstance.WriteDataToDataAccessService(grpcDataAccessServiceClient, dataAccessServiceWriteRequest, new IDataAccessServiceWriteResponse(), options1);
                            }
                            catch (Exception e)
                            {
                                Console.WriteLine("ERRO: {0}", e);
                                logNet.WriteError("设备编号4790的设备总览信息发送失败，错误原因 : " + e.ToString());
                            }
                            listWriteItem.Clear();

                            #endregion

                            stepNumber = 15;

                        
                    }
                        break;
                    case 15:
                        {

                            #region 读取并发送1000ms数据 点位名

                            //实例化发给IEC的 1000ms数据的点位名 结构体
                            var OneSecNameStruct = new OneSecPointNameStruct_IEC();

                            // 功能安全、生产统计、寿命管理 的点位名
                            keyenceClients.ReadPointName(Production_Data, ref OneSecNameStruct);
                            keyenceClients.ReadPointName(Function_Enable, ref OneSecNameStruct);
                            keyenceClients.ReadPointName(Life_Management, ref OneSecNameStruct);

                            //报警信息 的 点位名
                            var stringnumber = Alarm1.Length + Alarm2.Length + Alarm3.Length + Alarm4.Length;
                            var AlarmGroups = new List<OneSecInfoStruct_MC[]> { Alarm1, Alarm2, Alarm3, Alarm4 };
                            keyenceClients.ReadPointName(AlarmGroups, "Alarm", stringnumber, ref OneSecNameStruct);

                            //OEE 的点位名
                            stringnumber = OEE1.Length + OEE2.Length;
                            var OEEGroups = new List<OneSecInfoStruct_MC[]> { OEE1, OEE2 };
                            keyenceClients.ReadPointName(OEEGroups, "OEE", stringnumber, ref OneSecNameStruct);


                            //Grpc发送1000ms数据点位名结构体
                            listWriteItem.Clear();
                            try
                            {
                                listWriteItem.Add(grpcToolInstance.CreatWriteItem(nodeidDictionary1.GetValueOrDefault("OneSecNameStruct"), Arp.Type.Grpc.CoreType.CtStruct, OneSecNameStruct));
                                var writeItemsArray = listWriteItem.ToArray();
                                var dataAccessServiceWriteRequest = grpcToolInstance.ServiceWriteRequestAddDatas(writeItemsArray);
                                bool result = grpcToolInstance.WriteDataToDataAccessService(grpcDataAccessServiceClient, dataAccessServiceWriteRequest, new IDataAccessServiceWriteResponse(), options1);
                            }
                            catch (Exception e)
                            {
                                logNet.WriteError("[Grpc]", " 1000ms数据的点位名发送失败：" + e);
                                //Console.WriteLine("ERRO: {0}", e);
                            }

                            #endregion


                            #region 读取并发送加工工位数据 点位名

                            var ProcessStationNameStruct = new ProcessStationNameStruct_IEC();
                            List<StationInfoStruct_MC[]> StationDataStruct = new List<StationInfoStruct_MC[]>
                            { ZhuYeWei, JingZhiWei,FengZhuangWei};

                            keyenceClients.ReadPointName(StationDataStruct, ref ProcessStationNameStruct);

                            //Grpc发送1000ms数据点位名结构体
                            listWriteItem.Clear();
                            try
                            {
                                listWriteItem.Add(grpcToolInstance.CreatWriteItem(nodeidDictionary1.GetValueOrDefault("ProcessStationNameStruct"), Arp.Type.Grpc.CoreType.CtStruct, ProcessStationNameStruct));
                                var writeItemsArray = listWriteItem.ToArray();
                                var dataAccessServiceWriteRequest = grpcToolInstance.ServiceWriteRequestAddDatas(writeItemsArray);
                                bool result = grpcToolInstance.WriteDataToDataAccessService(grpcDataAccessServiceClient, dataAccessServiceWriteRequest, new IDataAccessServiceWriteResponse(), options1);
                            }
                            catch (Exception e)
                            {
                                logNet.WriteError("[Grpc]", " 加工工位的点位名发送失败：" + e);
                                //Console.WriteLine("ERRO: {0}", e);
                            }



                            #endregion



                            #region 读取并发送设备信息 （虚拟码 + 后工位码） 点位名

                            keyenceClients.ReadandSendStaionInfo(Battery_Memory, grpcToolInstance, nodeidDictionary1, grpcDataAccessServiceClient, options1);

                            #endregion

                            logNet.WriteInfo("点位名发送完毕");

                            stepNumber = 20;


                        }
                        break;




                    case 20:
                    {

                            #region MC连接

                            var i = 0;     //_mc[0]:4793
                            _mc[i] = new KeyenceMcNet(deviceInfoStruct1_IEC[0].strIPAddress, 5000);  //mc协议的端口号5000
                            var retConnect = _mc[i].ConnectServer();
                            //Console.WriteLine("num {0} connect: {1})!", i, retConnect.IsSuccess ? "success" : "fail");
                            logNet.WriteInfo("[MC]", "MC[0]连接：" + (retConnect.IsSuccess ? "成功" : "失败"));
                            logNet.WriteInfo("[MC]", "MC[0]连接设备的ip地址为：" + deviceInfoStruct1_IEC[0].strIPAddress);


                            i = 1;     //_mc[0]:4790
                            _mc[i] = new KeyenceMcNet(deviceInfoStruct2_IEC[0].strIPAddress, 5000);  //mc协议的端口号5000
                            retConnect = _mc[i].ConnectServer();
                            //Console.WriteLine("num {0} connect: {1})!", i, retConnect.IsSuccess ? "success" : "fail");
                            logNet.WriteInfo("[MC]", "MC[1]连接：" + (retConnect.IsSuccess ? "成功" : "失败"));
                            logNet.WriteInfo("[MC]", "MC[1]连接设备的ip地址为：" + deviceInfoStruct2_IEC[0].strIPAddress);


                            #endregion


                            stepNumber = 90;

                        break;

                    }

                    case 90:
                    {

                            //线程初始化      

                            #region 编号4793

                            //读设备信息
                            thr[0] = new Thread(() =>
                            {
                                var mc = _mc[0];
                                var nodeidDictionary = nodeidDictionary1;

                                var StationListInfo = new UDT_StationListInfo();
                                StationListInfo.iDataCount = (short)Battery_Memory.Length;

                                var listWriteItem = new List<WriteItem>();
                                WriteItem[] writeItems = new WriteItem[] { };


                                while (isThreadZeroRunning)
                                {
                                    TimeSpan start = new TimeSpan(DateTime.Now.Ticks);

                                    keyenceClients.ReadDeviceInfoStruct(Battery_Memory, mc, ref StationListInfo);
                                    keyenceClients.ReadDeviceInfoStruct(Battery_Clear, mc, ref StationListInfo);
                                    keyenceClients.ReadDeviceInfoStruct(Device_Enable, mc, ref StationListInfo);                                
                                    keyenceClients.ReadDeivceInfoStruct(BarCode, mc, ref allDataReadfromMC_4793, ref StationListInfo);
                                    keyenceClients.ReadDeivceInfoStruct(EarCode, mc, ref allDataReadfromMC_4793, ref StationListInfo);


                                    // Grpc发送数据给IEC                     
                                    try
                                    {
                                        listWriteItem.Add(grpcToolInstance.CreatWriteItem(nodeidDictionary["StationListInfo"], Arp.Type.Grpc.CoreType.CtStruct, StationListInfo));
                                        var writeItemsArray = listWriteItem.ToArray();
                                        var dataAccessServiceWriteRequest = grpcToolInstance.ServiceWriteRequestAddDatas(writeItemsArray);
                                        bool result = grpcToolInstance.WriteDataToDataAccessService(grpcDataAccessServiceClient, dataAccessServiceWriteRequest, new IDataAccessServiceWriteResponse(), options1);
                                    }
                                    catch (Exception e)
                                    {
                                        logNet.WriteError("[Grpc]", "设备信息数据发送失败：" + e);
                                        //Console.WriteLine("ERRO: {0}，{1}", e, nodeidDictionary.GetValueOrDefault(i.ToString()));
                                    }
                                    listWriteItem.Clear();



                                    TimeSpan end = new TimeSpan(DateTime.Now.Ticks);
                                    DateTime nowDisplay = DateTime.Now;
                                    TimeSpan dur =  (end - start).Duration();

                                    //logNet.WriteInfo("Thread ReadDeviceInfo read time : " + (dur.TotalMilliseconds).ToString());
                                    //Console.WriteLine("Thread ReadDeviceInfo read time:{0} read Duration:{1}", nowDisplay.ToString("yyyy-MM-dd HH:mm:ss:fff"), dur.TotalMilliseconds);

                                    if (dur.TotalMilliseconds < 100)
                                    {
                                        int sleepTime = 100 - (int)dur.TotalMilliseconds;
                                        Thread.Sleep(sleepTime);
                                    }
                                    else
                                    {
                                        //Console.WriteLine("Thread 4793 ReadDeviceInfo read time:{0} read Duration:{1}", nowDisplay.ToString("yyyy-MM-dd HH:mm:ss:fff"), dur.TotalMilliseconds);
                                        logNet.WriteInfo("No.4793 设备信息数据读取时间 : " + (dur.TotalMilliseconds).ToString());
                                    }
                                }


                            });

                            //读三大工位的信息  （硬编码）
                            thr[1] = new Thread(() =>
                            {
                                //硬编码的部分，读取36次减少为读取5次
                                short[] ZFarray1 = new short[8];
                                short[] ZFarray2 = new short[8];
                                short[] DMarray = new short[72];
                                short DM1814 = 0;
                                short DM1822 = 0;

                                var mc = _mc[0];
                                var nodeidDictionary = nodeidDictionary1;

                                var ProcessStationDataValue = new UDT_ProcessStationDataValue();   //实例化 加工工位采集值结构体

                                ProcessStationDataValue.iDataCount = 3;    // 一共四个加工工位

                                while (isThreadOneRunning)
                                {
                                    TimeSpan start = new TimeSpan(DateTime.Now.Ticks);

                                    OperateResult<short[]> ret_ZF1 = mc.ReadInt16("ZF200200", 8);
                                    if (ret_ZF1.IsSuccess)
                                    {
                                        Array.Copy(ret_ZF1.Content, ZFarray1, ret_ZF1.Content.Length);
                                    }
                                    else
                                    {
                                        logNet.WriteInfo("[MC]","ZF200200 数组读取失败");
                                        //Console.WriteLine(" ZF200200 Array Read failed");
                                    }

                                    OperateResult<short[]> ret_ZF2 = mc.ReadInt16("ZF20000", 8);
                                    if (ret_ZF2.IsSuccess)
                                    {
                                        Array.Copy(ret_ZF2.Content, ZFarray2, ret_ZF2.Content.Length);
                                    }
                                    else
                                    {
                                        logNet.WriteInfo("[MC]", "ZF20000 数组读取失败");
                                        // Console.WriteLine(" ZF20000 Array Read failed");
                                    }


                                    OperateResult<short[]> ret_DM = mc.ReadInt16("DM607", 72);
                                    if (ret_DM.IsSuccess)
                                    {
                                        Array.Copy(ret_DM.Content, DMarray, ret_DM.Content.Length);
                                    }
                                    else
                                    {
                                        logNet.WriteInfo("[MC]", "DM 数组读取失败");
                                        //Console.WriteLine(" DMArray Array Read failed");
                                    }

                                    OperateResult<short> ret_DM1814 = mc.ReadInt16("DM1814");
                                    if (ret_DM1814.IsSuccess)
                                    {
                                        DM1814 = ret_DM1814.Content;
                                    }
                                    else
                                    {
                                        logNet.WriteInfo("[MC]", "DM1814 读取失败");
                                        //Console.WriteLine(" DM1814 Array Read failed");
                                    }

                                    OperateResult<short> ret_DM1822= mc.ReadInt16("DM1822");
                                    if (ret_DM1822.IsSuccess)
                                    {
                                        DM1822 = ret_DM1822.Content;
                                    }
                                    else
                                    {
                                        logNet.WriteInfo("[MC]", "DM1822 读取失败");
                                        //Console.WriteLine(" DM1822 Array Read failed");
                                    }


                                    if (ret_ZF1.IsSuccess && ret_ZF2.IsSuccess && ret_DM.IsSuccess && ret_DM1814.IsSuccess && ret_DM1822.IsSuccess )
                                    {

                                        keyenceClients.SendStationData(ZhuYeWei, ref allDataReadfromMC_4793, ref ProcessStationDataValue, DMarray, 0);
                                        keyenceClients.SendStationData(JingZhiWei, ref allDataReadfromMC_4793, ref ProcessStationDataValue, DMarray, DM1814);
                                        keyenceClients.SendStationData(FengZhuangWei, ref allDataReadfromMC_4793, ref ProcessStationDataValue, ZFarray1, ZFarray2, DMarray, DM1822);


                                        //Grpc 发送加工工位数据采集值

                                        listWriteItem.Clear();
                                        try
                                        {
                                            listWriteItem.Add(grpcToolInstance.CreatWriteItem(nodeidDictionary["ProcessStationData"], Arp.Type.Grpc.CoreType.CtStruct, ProcessStationDataValue));
                                            var writeItemsArray = listWriteItem.ToArray();
                                            var dataAccessServiceWriteRequest = grpcToolInstance.ServiceWriteRequestAddDatas(writeItemsArray);
                                            bool result = grpcToolInstance.WriteDataToDataAccessService(grpcDataAccessServiceClient, dataAccessServiceWriteRequest, new IDataAccessServiceWriteResponse(), options1);
                                        }
                                        catch (Exception e)
                                        {
                                            logNet.WriteError("[Grpc]", "加工工位数据发送失败：" + e);

                                        }

                                    }
                                    else
                                    {

                                        logNet.WriteError("[MC]", "加工工位数据读取失败");

                                    }

                                    TimeSpan end = new TimeSpan(DateTime.Now.Ticks);
                                    DateTime nowDisplay = DateTime.Now;
                                    TimeSpan dur = (end - start).Duration();

                                    //logNet.WriteInfo("Thread ReadStationInfo read time : " + (dur.TotalMilliseconds).ToString());
                                    //Console.WriteLine("Thread ReadStationInfo read time:{0} read Duration:{1}", nowDisplay.ToString("yyyy-MM-dd HH:mm:ss:fff"), dur.TotalMilliseconds);

                                    if (dur.TotalMilliseconds < 100)
                                    {
                                        int sleepTime = 100 - (int)dur.TotalMilliseconds;
                                        Thread.Sleep(sleepTime);
                                    }
                                    else
                                    {
                                        //Console.WriteLine("No.4793 Thread ReadStationInfo read time:{0} read Duration:{1}", nowDisplay.ToString("yyyy-MM-dd HH:mm:ss:fff"), dur.TotalMilliseconds);
                                        logNet.WriteInfo("No.4793 加工工位数据读取时间 : " + (dur.TotalMilliseconds).ToString());
                                    }

                                }

                            });

                            //读1000ms数据
                            thr[2] = new Thread(() =>
                            {
                                var mc = _mc[0];
                                var nodeidDictionary = nodeidDictionary1;

                                var DeviceDataStruct = new DeviceDataStruct_IEC();
                            
                                var listWriteItem = new List<WriteItem>();
                                WriteItem[] writeItems = new WriteItem[] { };


                                while (isThreadTwoRunning)
                                {
                                    TimeSpan start = new TimeSpan(DateTime.Now.Ticks);

                                    #region 读取数据

                                    //读取并发送 生产统计、寿命管理、功能开关的数据
                                    keyenceClients.ReadOneSecData(Function_Enable, mc, ref allDataReadfromMC_4793, ref DeviceDataStruct);
                                    keyenceClients.ReadOneSecData(Production_Data, mc, ref allDataReadfromMC_4793, ref DeviceDataStruct);
                                    keyenceClients.ReadOneSecData(Life_Management, mc, ref allDataReadfromMC_4793, ref DeviceDataStruct);

                                    // 读取OEE数据
                                    bool[] OEE_part1 = keyenceClients.ReadOneSecConData(OEE1, mc);

                                    if (OEE_part1 != null)
                                    {
                                        Array.Copy(OEE_part1, 0, allDataReadfromMC_4793.OEEInfo1Value, 0, OEE_part1.Length);   //写入缓存区
                                    }

                                    bool[] OEE_part2 = keyenceClients.ReadOneSecConData(OEE2, mc);
                                    if (OEE_part2 != null)
                                    {
                                        Array.Copy(OEE_part2, 0, allDataReadfromMC_4793.OEEInfo2Value, 0, OEE_part2.Length);   //写入缓存区

                                    }

                                    //var OEEValue = new bool[OEE_part1.Length + OEE_part2.Length];  //与IEC对应 15
                             
                                    //将 OEE1和OEE2 拼成一个OEE数组

                                    if (OEE_part1 != null && OEE_part2! != null)
                                    {
                                        Array.Copy(OEE_part1, 0, DeviceDataStruct.Value_OEE, 0, OEE_part1.Length);
                                        Array.Copy(OEE_part2, 0, DeviceDataStruct.Value_OEE, OEE_part1.Length, OEE_part2.Length);
                                    }

                                    // 读取报警信号数据

                                    var alarmGroups = new List<bool[]> {
                                                keyenceClients.ReadOneSecDisData(Alarm1, mc, false),
                                                keyenceClients.ReadOneSecDisData(Alarm2, mc, false),
                                                keyenceClients.ReadOneSecDisData(Alarm3, mc, false),
                                                keyenceClients.ReadOneSecDisData(Alarm4, mc, true) };
                                    var ArrayIndex = 0;

                                    foreach (var alarmGroup in alarmGroups)
                                    {
                                        Array.Copy(alarmGroup, 0, DeviceDataStruct.Value_ALM, ArrayIndex, alarmGroup.Length);
                                        ArrayIndex += alarmGroup.Length;
                                    }

                                    //Grpc 发送1000ms数据采集值

                                    listWriteItem.Clear();

                                    try
                                    {
                                        listWriteItem.Add(grpcToolInstance.CreatWriteItem(nodeidDictionary["OneSecDataValue"], Arp.Type.Grpc.CoreType.CtStruct, DeviceDataStruct));
                                        var writeItemsArray = listWriteItem.ToArray();
                                        var dataAccessServiceWriteRequest = grpcToolInstance.ServiceWriteRequestAddDatas(writeItemsArray);
                                        bool result = grpcToolInstance.WriteDataToDataAccessService(grpcDataAccessServiceClient, dataAccessServiceWriteRequest, new IDataAccessServiceWriteResponse(), options1);
                                    }
                                    catch (Exception e)
                                    {
                                        //logNet.WriteError("[Grpc]", "OEE数据发送失败：" + e);
                                        Console.WriteLine("ERRO: {0}", e, nodeidDictionary.GetValueOrDefault("OneSecDataValue"));
                                    }



                                    #endregion


                                    TimeSpan end = new TimeSpan(DateTime.Now.Ticks);
                                    DateTime nowDisplay = DateTime.Now;
                                    TimeSpan dur = (end - start).Duration();

                                    //logNet.WriteInfo("Thread ReadOneSecInfo read time : " + (dur.TotalMilliseconds).ToString());
                                    //Console.WriteLine("Thread ReadOneSecInfo read time:{0} read Duration:{1}", nowDisplay.ToString("yyyy-MM-dd HH:mm:ss:fff"), dur.TotalMilliseconds);

                                    if (dur.TotalMilliseconds < 1000)
                                    {
                                        int sleepTime = 1000 - (int)dur.TotalMilliseconds;
                                        Thread.Sleep(sleepTime);
                                    }
                                    else
                                    {
                                        //Console.WriteLine("Thread 4793 ReadOneSecInfo read time:{0} read Duration:{1}", nowDisplay.ToString("yyyy-MM-dd HH:mm:ss:fff"), dur.TotalMilliseconds);
                                        logNet.WriteInfo("No.4793 1000ms数据读取时间 : " + (dur.TotalMilliseconds).ToString());
                                    }

                                }

                            });

                            #endregion



                            #region 编号4790

                            //读设备信息
                            thr[3] = new Thread(() =>
                            {
                                var mc = _mc[1];
                                var nodeidDictionary = nodeidDictionary2;

                                var StationListInfo = new UDT_StationListInfo();
                                StationListInfo.iDataCount = (short)Battery_Memory.Length;

                                var listWriteItem = new List<WriteItem>();
                                WriteItem[] writeItems = new WriteItem[] { };



                                while (isThreadThreeRunning)
                                {
                                    TimeSpan start = new TimeSpan(DateTime.Now.Ticks);

                                    keyenceClients.ReadDeviceInfoStruct(Battery_Memory, mc, ref StationListInfo);
                                    keyenceClients.ReadDeviceInfoStruct(Battery_Clear, mc, ref StationListInfo);
                                    keyenceClients.ReadDeviceInfoStruct(Device_Enable, mc, ref StationListInfo);
                                    keyenceClients.ReadDeivceInfoStruct(BarCode, mc, ref allDataReadfromMC_4790, ref StationListInfo);
                                    keyenceClients.ReadDeivceInfoStruct(EarCode, mc, ref allDataReadfromMC_4790, ref StationListInfo);


                                    // Grpc发送数据给IEC                     
                                    try
                                    {
                                        listWriteItem.Add(grpcToolInstance.CreatWriteItem(nodeidDictionary["StationListInfo"], Arp.Type.Grpc.CoreType.CtStruct, StationListInfo));
                                        var writeItemsArray = listWriteItem.ToArray();
                                        var dataAccessServiceWriteRequest = grpcToolInstance.ServiceWriteRequestAddDatas(writeItemsArray);
                                        bool result = grpcToolInstance.WriteDataToDataAccessService(grpcDataAccessServiceClient, dataAccessServiceWriteRequest, new IDataAccessServiceWriteResponse(), options1);
                                    }
                                    catch (Exception e)
                                    {
                                        logNet.WriteError("[Grpc]", "设备信息数据发送失败：" + e);
                                        //Console.WriteLine("ERRO: {0}，{1}", e, nodeidDictionary.GetValueOrDefault(i.ToString()));
                                    }
                                    listWriteItem.Clear();



                                    TimeSpan end = new TimeSpan(DateTime.Now.Ticks);
                                    DateTime nowDisplay = DateTime.Now;
                                    TimeSpan dur = (end - start).Duration();

                                    //logNet.WriteInfo("Thread ReadDeviceInfo read time : " + (dur.TotalMilliseconds).ToString());

                                    if (dur.TotalMilliseconds < 100)
                                    {
                                        int sleepTime = 100 - (int)dur.TotalMilliseconds;
                                        Thread.Sleep(sleepTime);
                                    }
                                    else
                                    {
                                        //Console.WriteLine("Thread 4790 ReadDeviceInfo read time:{0} read Duration:{1}", nowDisplay.ToString("yyyy-MM-dd HH:mm:ss:fff"), dur.TotalMilliseconds);
                                        logNet.WriteInfo("No.4790 设备信息读取时间 : " + (dur.TotalMilliseconds).ToString());
                                    }
                                }


                            });


                            //读三大工位的信息  （硬编码）
                            thr[4] = new Thread(() =>
                            {
                                //硬编码的部分，读取36次减少为读取5次
                                short[] ZFarray1 = new short[8];
                                short[] ZFarray2 = new short[8];
                                short[] DMarray = new short[72];
                                short DM1814 = 0;
                                short DM1822 = 0;

                                var mc = _mc[1];
                                var nodeidDictionary = nodeidDictionary2;

                                var ProcessStationDataValue = new UDT_ProcessStationDataValue();   //实例化 加工工位采集值结构体

                                ProcessStationDataValue.iDataCount = 3;    // 一共四个加工工位


                                while (isThreadFourRunning)
                                {
                                    TimeSpan start = new TimeSpan(DateTime.Now.Ticks);

                                    OperateResult<short[]> ret_ZF1 = mc.ReadInt16("ZF200200", 8);
                                    if (ret_ZF1.IsSuccess)
                                    {
                                        Array.Copy(ret_ZF1.Content, ZFarray1, ret_ZF1.Content.Length);
                                    }
                                    else
                                    {
                                        logNet.WriteInfo("[MC]", "ZF200200 数组读取失败");
                                        //Console.WriteLine(" ZF200200 Array Read failed");
                                    }

                                    OperateResult<short[]> ret_ZF2 = mc.ReadInt16("ZF20000", 8);
                                    if (ret_ZF2.IsSuccess)
                                    {
                                        Array.Copy(ret_ZF2.Content, ZFarray2, ret_ZF2.Content.Length);
                                    }
                                    else
                                    {
                                        logNet.WriteInfo("[MC]", "ZF20000 数组读取失败");
                                        // Console.WriteLine(" ZF20000 Array Read failed");
                                    }


                                    OperateResult<short[]> ret_DM = mc.ReadInt16("DM607", 72);
                                    if (ret_DM.IsSuccess)
                                    {
                                        Array.Copy(ret_DM.Content, DMarray, ret_DM.Content.Length);
                                    }
                                    else
                                    {
                                        logNet.WriteInfo("[MC]", "DM 数组读取失败");
                                        //Console.WriteLine(" DMArray Array Read failed");
                                    }

                                    OperateResult<short> ret_DM1814 = mc.ReadInt16("DM1814");
                                    if (ret_DM1814.IsSuccess)
                                    {
                                        DM1814 = ret_DM1814.Content;
                                    }
                                    else
                                    {
                                        logNet.WriteInfo("[MC]", "DM1814 读取失败");
                                        //Console.WriteLine(" DM1814 Array Read failed");
                                    }

                                    OperateResult<short> ret_DM1822 = mc.ReadInt16("DM1822");
                                    if (ret_DM1822.IsSuccess)
                                    {
                                        DM1822 = ret_DM1822.Content;
                                    }
                                    else
                                    {
                                        logNet.WriteInfo("[MC]", "DM1822 读取失败");
                                        //Console.WriteLine(" DM1822 Array Read failed");
                                    }


                                    if (ret_ZF1.IsSuccess && ret_ZF2.IsSuccess && ret_DM.IsSuccess && ret_DM1814.IsSuccess && ret_DM1822.IsSuccess)
                                    {

                                        keyenceClients.SendStationData(ZhuYeWei, ref allDataReadfromMC_4790, ref ProcessStationDataValue, DMarray, 0);
                                        keyenceClients.SendStationData(JingZhiWei, ref allDataReadfromMC_4790, ref ProcessStationDataValue, DMarray, DM1814);
                                        keyenceClients.SendStationData(FengZhuangWei, ref allDataReadfromMC_4790, ref ProcessStationDataValue, ZFarray1, ZFarray2, DMarray, DM1822);


                                        //Grpc 发送加工工位数据采集值

                                        listWriteItem.Clear();
                                        try
                                        {
                                            listWriteItem.Add(grpcToolInstance.CreatWriteItem(nodeidDictionary["ProcessStationData"], Arp.Type.Grpc.CoreType.CtStruct, ProcessStationDataValue));
                                            var writeItemsArray = listWriteItem.ToArray();
                                            var dataAccessServiceWriteRequest = grpcToolInstance.ServiceWriteRequestAddDatas(writeItemsArray);
                                            bool result = grpcToolInstance.WriteDataToDataAccessService(grpcDataAccessServiceClient, dataAccessServiceWriteRequest, new IDataAccessServiceWriteResponse(), options1);
                                        }
                                        catch (Exception e)
                                        {
                                            logNet.WriteError("[Grpc]", "加工工位数据发送失败：" + e);

                                        }

                                    }
                                    else
                                    {

                                        logNet.WriteError("[MC]", "加工工位数据读取失败");

                                    }

                                    TimeSpan end = new TimeSpan(DateTime.Now.Ticks);
                                    DateTime nowDisplay = DateTime.Now;
                                    TimeSpan dur = (end - start).Duration();

                                    //logNet.WriteInfo("Thread ReadStationInfo read time : " + (dur.TotalMilliseconds).ToString());

                                    if (dur.TotalMilliseconds < 100)
                                    {
                                        int sleepTime = 100 - (int)dur.TotalMilliseconds;
                                        Thread.Sleep(sleepTime);
                                    }
                                    else
                                    {
                                        //Console.WriteLine("Thread 4790 ReadStationInfo read time:{0} read Duration:{1}", nowDisplay.ToString("yyyy-MM-dd HH:mm:ss:fff"), dur.TotalMilliseconds);
                                        logNet.WriteInfo("No.4790 加工工位数据读取时间 : " + (dur.TotalMilliseconds).ToString());
                                    }

                                }

                            });


                            //读1000ms数据
                            thr[5] = new Thread(() =>
                            {
                                var mc = _mc[1];
                                var nodeidDictionary = nodeidDictionary2;

                                var DeviceDataStruct = new DeviceDataStruct_IEC();

                                var listWriteItem = new List<WriteItem>();
                                WriteItem[] writeItems = new WriteItem[] { };



                                while (isThreadFiveRunning)
                                {
                                    TimeSpan start = new TimeSpan(DateTime.Now.Ticks);

                                    #region 读取数据

                                    //读取并发送 生产统计、寿命管理、功能开关的数据
                                    keyenceClients.ReadOneSecData(Function_Enable, mc, ref allDataReadfromMC_4790, ref DeviceDataStruct);
                                    keyenceClients.ReadOneSecData(Production_Data, mc, ref allDataReadfromMC_4790, ref DeviceDataStruct);
                                    keyenceClients.ReadOneSecData(Life_Management, mc, ref allDataReadfromMC_4790, ref DeviceDataStruct);

                                    // 读取OEE数据
                                    bool[] OEE_part1 = keyenceClients.ReadOneSecConData(OEE1, mc);

                                    if (OEE_part1 != null)
                                    {
                                        Array.Copy(OEE_part1, 0, allDataReadfromMC_4790.OEEInfo1Value, 0, OEE_part1.Length);   //写入缓存区
                                    }

                                    bool[] OEE_part2 = keyenceClients.ReadOneSecConData(OEE2, mc);
                                    if (OEE_part2 != null)
                                    {
                                        Array.Copy(OEE_part2, 0, allDataReadfromMC_4790.OEEInfo2Value, 0, OEE_part2.Length);   //写入缓存区

                                    }

                                    //var OEEValue = new bool[OEE_part1.Length + OEE_part2.Length];  //与IEC对应 15

                                    //将 OEE1和OEE2 拼成一个OEE数组

                                    if (OEE_part1 != null && OEE_part2! != null)
                                    {
                                        Array.Copy(OEE_part1, 0, DeviceDataStruct.Value_OEE, 0, OEE_part1.Length);
                                        Array.Copy(OEE_part2, 0, DeviceDataStruct.Value_OEE, OEE_part1.Length, OEE_part2.Length);
                                    }

                                    // 读取报警信号数据

                                    var alarmGroups = new List<bool[]> {
                                            keyenceClients.ReadOneSecDisData(Alarm1, mc, false),
                                            keyenceClients.ReadOneSecDisData(Alarm2, mc, false),
                                            keyenceClients.ReadOneSecDisData(Alarm3, mc, false),
                                            keyenceClients.ReadOneSecDisData(Alarm4, mc, true) };
                                    var ArrayIndex = 0;

                                    foreach (var alarmGroup in alarmGroups)
                                    {
                                        Array.Copy(alarmGroup, 0, DeviceDataStruct.Value_ALM, ArrayIndex, alarmGroup.Length);
                                        ArrayIndex += alarmGroup.Length;
                                    }

                                    //Grpc 发送1000ms数据采集值

                                    listWriteItem.Clear();

                                    try
                                    {
                                        listWriteItem.Add(grpcToolInstance.CreatWriteItem(nodeidDictionary["OneSecDataValue"], Arp.Type.Grpc.CoreType.CtStruct, DeviceDataStruct));
                                        var writeItemsArray = listWriteItem.ToArray();
                                        var dataAccessServiceWriteRequest = grpcToolInstance.ServiceWriteRequestAddDatas(writeItemsArray);
                                        bool result = grpcToolInstance.WriteDataToDataAccessService(grpcDataAccessServiceClient, dataAccessServiceWriteRequest, new IDataAccessServiceWriteResponse(), options1);
                                    }
                                    catch (Exception e)
                                    {
                                        //logNet.WriteError("[Grpc]", "OEE数据发送失败：" + e);
                                        Console.WriteLine("ERRO: {0}", e, nodeidDictionary.GetValueOrDefault("OneSecDataValue"));
                                    }



                                    #endregion



                                    TimeSpan end = new TimeSpan(DateTime.Now.Ticks);
                                    DateTime nowDisplay = DateTime.Now;
                                    TimeSpan dur = (end - start).Duration();

                                    //logNet.WriteInfo("Thread ReadOneSecInfo read time : " + (dur.TotalMilliseconds).ToString());

                                    if (dur.TotalMilliseconds < 1000)
                                    {
                                        int sleepTime = 1000 - (int)dur.TotalMilliseconds;
                                        Thread.Sleep(sleepTime);
                                    }
                                    else
                                    {
                                        logNet.WriteInfo("No.4790 1000ms数据读取时间 : " + (dur.TotalMilliseconds).ToString());                      
                                        //Console.WriteLine("Thread 4790 ReadOneSecInfo read time:{0} read Duration:{1}", nowDisplay.ToString("yyyy-MM-dd HH:mm:ss:fff"), dur.TotalMilliseconds);
                                    }

                                }

                            });


                            #endregion

          

                            stepNumber = 100;

                            break;

                    }

                    case 100:
                    {
                            #region 开启线程

                        if (thr[0].ThreadState == ThreadState.Unstarted && thr[1].ThreadState == ThreadState.Unstarted && thr[2].ThreadState == ThreadState.Unstarted
                              &&thr[3].ThreadState == ThreadState.Unstarted && thr[4].ThreadState == ThreadState.Unstarted && thr[5].ThreadState == ThreadState.Unstarted  )
                        {   
                            try
                                {
                                    //编号 4793
                                    isThreadZeroRunning = true;
                                    thr[0].Start();  // 读设备信息

                                     isThreadOneRunning = true;
                                    thr[1].Start(); //读三大工位信息

                                    isThreadTwoRunning = true;
                                    thr[2].Start();  //读1000ms 数据

                                    ////编号4790
                                    isThreadThreeRunning = true;
                                    thr[3].Start();  // 读1000ms数据

                                    isThreadFourRunning = true;
                                    thr[4].Start(); //读三大工位信息

                                    isThreadFiveRunning = true;
                                    thr[5].Start();  //读设备信息


                                    //APP Status ： running
                                    listWriteItem.Clear();
                                    listWriteItem.Add(grpcToolInstance.CreatWriteItem(nodeidDictionary1["AppStatus"], Arp.Type.Grpc.CoreType.CtInt32, 1));
                                    if (grpcToolInstance.WriteDataToDataAccessService(grpcDataAccessServiceClient, grpcToolInstance.ServiceWriteRequestAddDatas(listWriteItem.ToArray()), new IDataAccessServiceWriteResponse(), options1))
                                    {
                                        logNet.WriteInfo("[Grpc]", "AppStatus 写入IEC成功");
                                        //Console.WriteLine("{0}      AppStatus写入IEC: success", DateTime.Now);
                                    }
                                    else
                                    {
                                        //Console.WriteLine("{0}      AppStatus写入IEC: fail", DateTime.Now);
                                        logNet.WriteError("[Grpc]", "AppStatus 写入IEC失败");
                                    }


                                }
                            catch
                            {
                                Console.WriteLine("Thread quit");

                                //APP Status ： Error
                                listWriteItem.Clear();
                                listWriteItem.Add(grpcToolInstance.CreatWriteItem(nodeidDictionary1["AppStatus"], Arp.Type.Grpc.CoreType.CtInt32, -1));
                                if (grpcToolInstance.WriteDataToDataAccessService(grpcDataAccessServiceClient, grpcToolInstance.ServiceWriteRequestAddDatas(listWriteItem.ToArray()), new IDataAccessServiceWriteResponse(), options1))
                                {
                                    logNet.WriteInfo("[Grpc]", "AppStatus 写入IEC成功");
                                    //Console.WriteLine("{0}      AppStatus写入IEC: success", DateTime.Now);
                                }
                                else
                                {
                                    //Console.WriteLine("{0}      AppStatus写入IEC: fail", DateTime.Now);
                                    logNet.WriteError("[Grpc]", "AppStatus 写入IEC失败");
                                }

                            }                                             
                        }
                            #endregion


                            #region IEC发送触发信号，重新读取Excel

                            dataAccessServiceReadSingleRequest = new IDataAccessServiceReadSingleRequest();
                            dataAccessServiceReadSingleRequest.PortName = nodeidDictionary1["Switch_ReadExcelFile"];
                            if (grpcToolInstance.ReadSingleDataToDataAccessService(grpcDataAccessServiceClient, dataAccessServiceReadSingleRequest, new IDataAccessServiceReadSingleResponse(), options1).BoolValue)
                            {
                                //复位信号点:Switch_WriteExcelFile                               
                                listWriteItem.Clear();
                                listWriteItem.Add(grpcToolInstance.CreatWriteItem(nodeidDictionary1["Switch_ReadExcelFile"], Arp.Type.Grpc.CoreType.CtBoolean, false)); //Write Data to DataAccessService                                 
                                if (grpcToolInstance.WriteDataToDataAccessService(grpcDataAccessServiceClient, grpcToolInstance.ServiceWriteRequestAddDatas(listWriteItem.ToArray()), new IDataAccessServiceWriteResponse(), options1))
                                {
                                    //Console.WriteLine("{0}      Switch_ReadExcelFile写入IEC: success", DateTime.Now);
                                    logNet.WriteInfo("[Grpc]", "Switch_ReadExcelFile 写入IEC成功");
                                }
                                else
                                {
                                    //Console.WriteLine("{0}      Switch_ReadExcelFile写入IEC: fail", DateTime.Now);
                                    logNet.WriteError("[Grpc]", "Switch_ReadExcelFile 写入IEC失败");
                                }


                                //停止线程
                                isThreadZeroRunning = false;
                                isThreadOneRunning = false;
                                isThreadTwoRunning = false;
                                isThreadThreeRunning = false;
                                isThreadFourRunning = false;
                                isThreadFiveRunning = false;

                                for (int i = 0; i < clientNum; i++)
                                {
                                    _mc[i].ConnectClose();
                                    //Console.WriteLine(" MC {0} Connect closed", i);
                                    logNet.WriteInfo("[MC]", "MC连接断开" + i.ToString());
                                }

                                Thread.Sleep(1000);//等待线程退出

                                stepNumber = 6;
                            }

                            #endregion


                            #region 检测PLCnext和Keyence PLC之间的连接

                            for (int i = 0; i < clientNum; i++)
                            {
                                IPStatus iPStatus;
                                iPStatus = _mc[i].IpAddressPing();  //判断与PLC的物理连接状态

                                string[] plcErrors = {

                                                        "Ping Keyence PLC 4793 failed",
                                                        "Ping Keyence PLC 4790 failed"

                                                        };

                                if (iPStatus != 0)
                                {
                                    logNet.WriteError("[MC]", plcErrors[i]);

                                    //APP Status ： Error
                                    listWriteItem.Clear();
                                    listWriteItem.Add(grpcToolInstance.CreatWriteItem(nodeidDictionary1["AppStatus"], Arp.Type.Grpc.CoreType.CtInt32, -2));
                                    if (grpcToolInstance.WriteDataToDataAccessService(grpcDataAccessServiceClient, grpcToolInstance.ServiceWriteRequestAddDatas(listWriteItem.ToArray()), new IDataAccessServiceWriteResponse(), options1))
                                    {
                                        logNet.WriteInfo("[Grpc]", "AppStatus 写入IEC成功");
                                        //Console.WriteLine("{0}      AppStatus写入IEC: success", DateTime.Now);
                                    }
                                    else
                                    {
                                        //Console.WriteLine("{0}      AppStatus写入IEC: fail", DateTime.Now);
                                        logNet.WriteError("[Grpc]", "AppStatus 写入IEC失败");
                                    }


                                }

                            }

                            #endregion



                            #region IEC发送触发信号,将采集值写入Excel

                            dataAccessServiceReadSingleRequest = new IDataAccessServiceReadSingleRequest();
                            dataAccessServiceReadSingleRequest.PortName = nodeidDictionary1["Switch_WriteExcelFile"];
                            if (grpcToolInstance.ReadSingleDataToDataAccessService(grpcDataAccessServiceClient, dataAccessServiceReadSingleRequest, new IDataAccessServiceReadSingleResponse(), options1).BoolValue)
                            {
                                //复位信号点: Switch_WriteExcelFile
                                listWriteItem.Clear();
                                listWriteItem.Add(grpcToolInstance.CreatWriteItem(nodeidDictionary1["Switch_WriteExcelFile"], Arp.Type.Grpc.CoreType.CtBoolean, false)); //Write Data to DataAccessService                                 
                                if (grpcToolInstance.WriteDataToDataAccessService(grpcDataAccessServiceClient, grpcToolInstance.ServiceWriteRequestAddDatas(listWriteItem.ToArray()), new IDataAccessServiceWriteResponse(), options1))
                                {
                                    //Console.WriteLine("{0}      Switch_WriteExcelFile: success", DateTime.Now);
                                    logNet.WriteInfo("[Grpc]", "Switch_WriteExcelFile 写入IEC成功");
                                }
                                else
                                {
                                    //Console.WriteLine("{0}      Switch_WriteExcelFile: fail", DateTime.Now);
                                    logNet.WriteError("[Grpc]", "Switch_WriteExcelFile 写入IEC失败");
                                }

                                //将读取的值写入Excel 
                                thr[6] = new Thread(() =>
                                {

                                    var ExcelPath1 = "/opt/plcnext/apps/ZYJData(4793).xlsx";
                                    var ExcelPath2 = "/opt/plcnext/apps/ZYJData(4790).xlsx";

                                    //var ExcelPath1 = Directory.GetCurrentDirectory() + "\\ZYJData(4793).xlsx";
                                    //var ExcelPath2 = Directory.GetCurrentDirectory() + "\\ZYJData(4790).xlsx";

                                    //将数据缓存区的值赋给临时变量
                                    var allDataReadfromMC_temp_4793 = allDataReadfromMC_4793;
                                    var allDataReadfromMC_temp_4790 = allDataReadfromMC_4790;



                                    #region 将数据缓存区的值写入Excel(4793)

                                    try
                                    {
                                        var result = readExcel.setExcelCellValue(ExcelPath1, "设备信息", "电芯条码地址采集值", allDataReadfromMC_temp_4793.BarCode);
                                        logNet.WriteInfo("WriteData", "编号4793 电芯条码地址采集值写入Excel: " + (result ? "成功" : "失败"));
                                    }
                                    catch (Exception e)
                                    {
                                        logNet.WriteError("WriteData", "编号4793 电芯条码地址采集值写入Excel失败原因: " + e);
                                    }

                                    try
                                    {
                                        var result = readExcel.setExcelCellValue(ExcelPath1, "设备信息", "极耳码地址采集值", allDataReadfromMC_temp_4793.EarCode);
                                        logNet.WriteInfo("WriteData", "编号4793 极耳码地址采集值写入Excel: " + (result ? "成功" : "失败"));
                                    }
                                    catch (Exception e)
                                    {
                                        logNet.WriteError("WriteData", "编号4793 极耳码地址采集值写入Excel失败原因: " + e);
                                    }



                                    try
                                    {
                                        var result = readExcel.setExcelCellValue(ExcelPath1, "加工工位（注液位）", "采集值", allDataReadfromMC_temp_4793.ZhuYeWeiValue);
                                        logNet.WriteInfo("WriteData", "编号4793 加工工位（注液位）采集值写入Excel: " + (result ? "成功" : "失败"));
                                    }
                                    catch (Exception e)
                                    {
                                        logNet.WriteError("WriteData", "编号4793 加工工位（注液位）采集值写入Excel失败原因: " + e);

                                    }

                                    try
                                    {
                                        var result = readExcel.setExcelCellValue(ExcelPath1, "加工工位（静置位）", "采集值", allDataReadfromMC_temp_4793.JingZhiWeiValue);
                                        logNet.WriteInfo("WriteData", "编号4793 加工工位（静置位）采集值写入Excel: " + (result ? "成功" : "失败"));
                                    }
                                    catch (Exception e)
                                    {
                                        logNet.WriteError("WriteData", "编号4793 加工工位（静置位）采集值写入Excel失败原因: " + e);
                                    }

                                    try
                                    {
                                        var result = readExcel.setExcelCellValue(ExcelPath1, "加工工位（封装位）", "采集值", allDataReadfromMC_temp_4793.FengZhuangValue);
                                        logNet.WriteInfo("WriteData", "编号4793 加工工位（封装位）采集值写入Excel: " + (result ? "成功" : "失败"));
                                    }
                                    catch (Exception e)
                                    {
                                        logNet.WriteError("WriteData", "编号4793 加工工位（封装位）采集值写入Excel失败原因: " + e);
                                    }


                                    try
                                    {
                                        var result = readExcel.setExcelCellValue(ExcelPath1, "OEE(1)", "采集值", allDataReadfromMC_temp_4793.OEEInfo1Value);
                                        logNet.WriteInfo("WriteData", "编号4793 OEE(1)采集值写入Excel: " + (result ? "成功" : "失败"));
                                    }
                                    catch (Exception e)
                                    {
                                        logNet.WriteError("WriteData", "编号4793 OEE(1)采集值写入Excel失败原因: " + e);
                                    }

                                    try
                                    {
                                        var result = readExcel.setExcelCellValue(ExcelPath1, "OEE(2)", "采集值", allDataReadfromMC_temp_4793.OEEInfo2Value);
                                        logNet.WriteInfo("WriteData", "编号4793 OEE(2)采集值写入Excel: " + (result ? "成功" : "失败"));
                                    }
                                    catch (Exception e)
                                    {
                                        logNet.WriteError("WriteData", "编号4793 OEE(2)采集值写入Excel失败原因: " + e);
                                    }

                                    try
                                    {
                                        var result = readExcel.setExcelCellValue(ExcelPath1, "功能开关", "采集值", allDataReadfromMC_temp_4793.FunctionEnableValue);
                                        logNet.WriteInfo("WriteData", "编号4793 功能开关采集值写入Excel: " + (result ? "成功" : "失败"));
                                    }
                                    catch (Exception e)
                                    {
                                        logNet.WriteError("WriteData", "编号4793 功能开关采集值写入Excel失败原因: " + e);
                                    }

                                    try
                                    {
                                        var result = readExcel.setExcelCellValue(ExcelPath1, "生产统计", "采集值", allDataReadfromMC_temp_4793.ProductionDataValue);
                                        logNet.WriteInfo("WriteData", "编号4793 生产统计采集值写入Excel: " + (result ? "成功" : "失败"));
                                    }
                                    catch (Exception e)
                                    {
                                        logNet.WriteError("WriteData", "编号4793 生产统计采集值写入Excel失败原因: " + e);
                                    }

                                    try
                                    {
                                        var result = readExcel.setExcelCellValue(ExcelPath1, "寿命管理", "采集值", allDataReadfromMC_temp_4793.LifeManagementValue);
                                        logNet.WriteInfo("WriteData", "编号4793 寿命管理采集值写入Excel: " + (result ? "成功" : "失败"));
                                    }
                                    catch (Exception e)
                                    {
                                        logNet.WriteError("WriteData", "编号4793 寿命管理采集值写入Excel失败原因: " + e);
                                    }

                                    #endregion

                                    #region 将数据缓存区的值写入Excel(4790)

                                    try
                                    {
                                        var result = readExcel.setExcelCellValue(ExcelPath2, "设备信息", "电芯条码地址采集值", allDataReadfromMC_temp_4790.BarCode);
                                        logNet.WriteInfo("WriteData", "编号4790 电芯条码地址采集值写入Excel: " + (result ? "成功" : "失败"));
                                    }
                                    catch (Exception e)
                                    {
                                        logNet.WriteError("WriteData", "编号4790 电芯条码地址采集值写入Excel失败原因: " + e);
                                    }

                                    try
                                    {
                                        var result = readExcel.setExcelCellValue(ExcelPath2, "设备信息", "极耳码地址采集值", allDataReadfromMC_temp_4790.EarCode);
                                        logNet.WriteInfo("WriteData", "编号4790 极耳码地址采集值写入Excel: " + (result ? "成功" : "失败"));
                                    }
                                    catch (Exception e)
                                    {
                                        logNet.WriteError("WriteData", "编号4790 极耳码地址采集值写入Excel失败原因: " + e);
                                    }



                                    try
                                    {
                                        var result = readExcel.setExcelCellValue(ExcelPath2, "加工工位（注液位）", "采集值", allDataReadfromMC_temp_4790.ZhuYeWeiValue);
                                        logNet.WriteInfo("WriteData", "编号4790 加工工位（注液位）采集值写入Excel: " + (result ? "成功" : "失败"));
                                    }
                                    catch (Exception e)
                                    {
                                        logNet.WriteError("WriteData", "编号4790 加工工位（注液位）采集值写入Excel失败原因: " + e);

                                    }

                                    try
                                    {
                                        var result = readExcel.setExcelCellValue(ExcelPath2, "加工工位（静置位）", "采集值", allDataReadfromMC_temp_4790.JingZhiWeiValue);
                                        logNet.WriteInfo("WriteData", "编号4790 加工工位（静置位）采集值写入Excel: " + (result ? "成功" : "失败"));
                                    }
                                    catch (Exception e)
                                    {
                                        logNet.WriteError("WriteData", "编号4790 加工工位（静置位）采集值写入Excel失败原因: " + e);
                                    }

                                    try
                                    {
                                        var result = readExcel.setExcelCellValue(ExcelPath2, "加工工位（封装位）", "采集值", allDataReadfromMC_temp_4790.FengZhuangValue);
                                        logNet.WriteInfo("WriteData", "编号4790 加工工位（封装位）采集值写入Excel: " + (result ? "成功" : "失败"));
                                    }
                                    catch (Exception e)
                                    {
                                        logNet.WriteError("WriteData", "编号4790 加工工位（封装位）采集值写入Excel失败原因: " + e);
                                    }


                                    try
                                    {
                                        var result = readExcel.setExcelCellValue(ExcelPath2, "OEE(1)", "采集值", allDataReadfromMC_temp_4790.OEEInfo1Value);
                                        logNet.WriteInfo("WriteData", "编号4790 OEE(1)采集值写入Excel: " + (result ? "成功" : "失败"));
                                    }
                                    catch (Exception e)
                                    {
                                        logNet.WriteError("WriteData", "编号4790 OEE(1)采集值写入Excel失败原因: " + e);
                                    }

                                    try
                                    {
                                        var result = readExcel.setExcelCellValue(ExcelPath2, "OEE(2)", "采集值", allDataReadfromMC_temp_4790.OEEInfo2Value);
                                        logNet.WriteInfo("WriteData", "编号4790 OEE(2)采集值写入Excel: " + (result ? "成功" : "失败"));
                                    }
                                    catch (Exception e)
                                    {
                                        logNet.WriteError("WriteData", "编号4790 OEE(2)采集值写入Excel失败原因: " + e);
                                    }

                                    try
                                    {
                                        var result = readExcel.setExcelCellValue(ExcelPath2, "功能开关", "采集值", allDataReadfromMC_temp_4790.FunctionEnableValue);
                                        logNet.WriteInfo("WriteData", "编号4790 功能开关采集值写入Excel: " + (result ? "成功" : "失败"));
                                    }
                                    catch (Exception e)
                                    {
                                        logNet.WriteError("WriteData", "编号4790 功能开关采集值写入Excel失败原因: " + e);
                                    }

                                    try
                                    {
                                        var result = readExcel.setExcelCellValue(ExcelPath2, "生产统计", "采集值", allDataReadfromMC_temp_4790.ProductionDataValue);
                                        logNet.WriteInfo("WriteData", "编号4790 生产统计采集值写入Excel: " + (result ? "成功" : "失败"));
                                    }
                                    catch (Exception e)
                                    {
                                        logNet.WriteError("WriteData", "编号4790 生产统计采集值写入Excel失败原因: " + e);
                                    }

                                    try
                                    {
                                        var result = readExcel.setExcelCellValue(ExcelPath2, "寿命管理", "采集值", allDataReadfromMC_temp_4790.LifeManagementValue);
                                        logNet.WriteInfo("WriteData", "编号4790 寿命管理采集值写入Excel: " + (result ? "成功" : "失败"));
                                    }
                                    catch (Exception e)
                                    {
                                        logNet.WriteError("WriteData", "编号4790 寿命管理采集值写入Excel失败原因: " + e);
                                    }

                                    #endregion



                                    //给IEC写入 采集值写入成功的信号
                                    var tempFlag_finishWriteExcelFile = true;

                                    listWriteItem.Clear();
                                    listWriteItem.Add(grpcToolInstance.CreatWriteItem(nodeidDictionary1["flag_finishWriteExcelFile"], Arp.Type.Grpc.CoreType.CtBoolean, tempFlag_finishWriteExcelFile));
                                    if (grpcToolInstance.WriteDataToDataAccessService(grpcDataAccessServiceClient, grpcToolInstance.ServiceWriteRequestAddDatas(listWriteItem.ToArray()), new IDataAccessServiceWriteResponse(), options1))
                                    {
                                        //Console.WriteLine("{0}      flag_finishWriteExcelFile写入IEC: success", DateTime.Now);
                                        logNet.WriteInfo("[Grpc]", "flag_finishWriteExcelFile 写入IEC成功");
                                    }
                                    else
                                    {
                                        //Console.WriteLine("{0}      flag_finishWriteExcelFile写入IEC: fail", DateTime.Now);
                                        logNet.WriteError("[Grpc]", "flag_finishWriteExcelFile 写入IEC失败");
                                    }

                                    IecTriggersNumber = 0;  //为了防止IEC连续两次赋值true

                                });

                                IecTriggersNumber++;

                                if (IecTriggersNumber == 1)
                                {
                                    thr[6].Start();
                                }

                            }

                            #endregion


                            Thread.Sleep(1000);
                       
                        break;
                    }



                    case 1000:      //异常处理
                                    //信号复位
                                    //CIP连接断了


                        break;

                    case 10000:      //复位处理

                        break;


                }
            }
        }
    }
}

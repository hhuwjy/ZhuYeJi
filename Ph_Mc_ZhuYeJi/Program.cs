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
        //const string logsFile = ("/opt/plcnext/apps/ZhuYeAppLogs.txt");
        const string logsFile = "D:\\2024\\Work\\12-冠宇数采项目\\ReadFromStructArray\\ZhuYeJi_MC";
        public static ILogNet logNet = new LogNetSingle(logsFile);

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
        public static Dictionary<string, string> nodeidDictionary;

        //读取Excel用
        static ReadExcel readExcel = new ReadExcel();

        //MC Client实例化 
        public static KeyenceComm keyenceClients = new KeyenceComm();
        static int clientNum = 2;  //一个EPC对应采集两个基恩士的数据（点表相同）  上位链路+MC协议，同时在线加起来不能超过15台
        public static KeyenceMcNet[] _mc = new KeyenceMcNet[clientNum];
        
        
        //PLC IP Address地址
        public static List<string> plcIpAddresses = new List<string>();

        //创建线程
        static Thread[] thr;

        #region 从Excel解析来的数据实例化

        // 设备信息（47个工位）

        static DeviceInfoConSturct_MC[] Battery_Memory;
        static DeviceInfoConSturct_MC[] Battery_Clear;
        static DeviceInfoConSturct_MC[] Device_Enable;

        // 设备总览
        static DeviceInfoStruct_IEC[] deviceInfoStruct_IEC;

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


        // 时间变量
        public static DateTime nowDisplay = DateTime.Now;



        static void Main(string[] args)
        {
            int stepNumber = 10;

            while (true)
            {
                switch (stepNumber)
                {
                    case 10:

                        /// <summary>
                        /// 执行初始化
                        /// </summary>


                        #region 读取Excel 

                        logNet.WriteError(nowDisplay.ToString("yyyy-MM-dd HH:mm:ss:fff") + "App Start");

                        string excelFilePath = Directory.GetCurrentDirectory() + "\\ZYJData.xlsx";     //PC端测试路径
                        //string excelFilePath = "/opt/plcnext/apps/ZYJData.xlsx";                         //EPC存放路径

                        XSSFWorkbook excelWorkbook = readExcel.connectExcel(excelFilePath);

                        Console.WriteLine("ExcelWorkbook read {0}", excelWorkbook != null ? "success" : "fail");
                        logNet.WriteError(nowDisplay.ToString("yyyy-MM-dd HH:mm:ss:fff") + "  :ExcelWorkbook read ", excelWorkbook != null ? "success" : "fail");

                        #endregion


                        #region 从xml获取nodeid，Grpc发送到对应变量时使用，注意xml中的别名要和对应类的属性名一致 
                        try
                        {
                            const string filePath = "/opt/plcnext/apps/GrpcSubscribeNodes.xml";             //EPC中存放的路径  
                                                                                                            //const string filePath = "D:\\2024\\Work\\12-冠宇数采项目\\ReadFromStructArray\\FengZhuang_EIP\\Ph_CipComm_FengZhuang\\GrpcSubscribeNodes\\GrpcSubscribeNodes.xml";  //PC中存放的路径 

                            nodeidDictionary = grpcToolInstance.getNodeIdDictionary(filePath);  //将xml中的值写入字典中
                            logNet.WriteError(nowDisplay.ToString("yyyy-MM-dd HH:mm:ss:fff") + "  :NodeID read successfully");

                        }
                        catch (Exception e)
                        {
                            logNet.WriteError("Error:" + e);
                            logNet.WriteError(nowDisplay.ToString("yyyy-MM-dd HH:mm:ss:fff") + "  :NodeID read failed");
                        }

                        #endregion


                        #region 将readExcel变量中的值，写入对应的实例化结构体中

                        // 单独设备总览表 + 发送
                        deviceInfoStruct_IEC = readExcel.ReadDeviceInfo_Excel(excelWorkbook, "封装设备总览");

                        //设备信息
                        Battery_Memory = readExcel.ReadOneDeviceInfoConSturctInfo_Excel(excelWorkbook, "设备信息",3);
                        Battery_Clear = readExcel.ReadOneDeviceInfoConSturctInfo_Excel(excelWorkbook, "设备信息", 4);
                        Device_Enable = readExcel.ReadOneDeviceInfoConSturctInfo_Excel(excelWorkbook, "设备信息", 5);

                        //1000ms数据
                        Function_Enable = readExcel.ReadOneSecInfo_Excel(excelWorkbook, "功能开关");
                        Production_Data = readExcel.ReadOneSecInfo_Excel(excelWorkbook, "生产统计");

                        OEE1 = readExcel.ReadOneSecInfo_Excel(excelWorkbook, "OEE(1)");
                        OEE2 = readExcel.ReadOneSecInfo_Excel(excelWorkbook, "OEE(2)");
                        Alarm1 = readExcel.ReadOneSecInfo_Excel(excelWorkbook, "报警信号1");
                        Alarm2 = readExcel.ReadOneSecInfo_Excel(excelWorkbook, "报警信号2");
                        Alarm3 = readExcel.ReadOneSecInfo_Excel(excelWorkbook, "报警信号3");
                        Alarm4 = readExcel.ReadOneSecInfo_Excel(excelWorkbook, "报警信号4");

                        //三大工位
                        ZhuYeWei = readExcel.ReadStationInfo_Excel(excelWorkbook, "加工工位（注液位）");
                        JingZhiWei = readExcel.ReadStationInfo_Excel(excelWorkbook, "加工工位（静置位）");
                        FengZhuangWei = readExcel.ReadStationInfo_Excel(excelWorkbook, "加工工位（封装位）");

                        #endregion


                        break;



                    case 20:

                        #region MC连接

                        for(int i=0;i<clientNum;i++)
                        {
                            _mc[i] = new KeyenceMcNet(deviceInfoStruct_IEC[i].strIPAddress,5000);  //mc协议的端口号5000
                            var retConnect = _mc[i].ConnectServer();
                            Console.WriteLine("num {0} connect: {1})!", i, retConnect.IsSuccess ? "success" : "fail");
                            logNet.WriteInfo("num " + i.ToString() + (retConnect.IsSuccess ? "success" : "fail"));
                        }

                        #endregion

                        break;


                    case 90:

                        #region 线程初始化（先读一个基恩士的）

                        //读第一个基恩士的设备信息
                        thr[0] = new Thread(() =>
                        {
                            var mc = _mc[0];

                            int NumberOfStation = 48;      //取1-47号工位

                            StringBuilder[] sbBatteryMemory = new StringBuilder[NumberOfStation];
                            StringBuilder[] sbBatteryClear = new StringBuilder[NumberOfStation];                          
                            StringBuilder[] sbDeviceEnable = new StringBuilder[NumberOfStation];                         

                            //StringBuilder数组初始化
                            for (int i = 0; i < NumberOfStation; i++)
                            {
                                sbBatteryMemory[i] = new StringBuilder();
                                sbBatteryClear[i] = new StringBuilder();
                                sbDeviceEnable[i] = new StringBuilder();
                            }

                            var listWriteItem = new List<WriteItem>();
                            WriteItem[] writeItems = new WriteItem[] { };

                            stringStruct[] sendStringtoIEC = new stringStruct[47];

                            while(true)
                            {
                                TimeSpan start = new TimeSpan(DateTime.Now.Ticks);

                                //清空数据缓存区
                                for (int i = 0; i < NumberOfStation; i++)
                                {
                                    sbBatteryMemory[i].Clear();
                                    sbBatteryClear[i].Clear();
                                    sbDeviceEnable[i].Clear();
                                }

                                keyenceClients.ReadDeviceInfo(Battery_Memory, mc, sbBatteryMemory);
                                keyenceClients.ReadDeviceInfo(Battery_Clear, mc, sbBatteryClear);
                                keyenceClients.ReadDeviceInfo(Device_Enable, mc, sbDeviceEnable);


                                //整合到47个string中 舍弃下标0 并行发送数据                         
                                for (int i = 1; i < NumberOfStation; i++)
                                {
                                    StringBuilder combinedString = new StringBuilder();   //每一行工位都是一个string

                                    combinedString.Append(sbBatteryMemory[i]);
                                    combinedString.Append(sbBatteryClear[i]);

                                    if (sbDeviceEnable[i].Length == 0)
                                    {
                                        combinedString.Append(sbDeviceEnable[i] + " ,");

                                    }
                                    else
                                    {
                                        combinedString.Append(sbDeviceEnable[i]);
                                    }

                                    sendStringtoIEC[i - 1].str = combinedString.ToString(); ; //整合到结构体数组中 （一把子把47个工位的数据发送给IEC）

                                    #region Grpc发送数据给IEC

                                    if (i == 47)
                                    {
                                        try
                                        {
                                            listWriteItem.Add(grpcToolInstance.CreatWriteItem(nodeidDictionary["Station"], Arp.Type.Grpc.CoreType.CtArray, sendStringtoIEC));
                                            var writeItemsArray = listWriteItem.ToArray();
                                            var dataAccessServiceWriteRequest = grpcToolInstance.ServiceWriteRequestAddDatas(writeItemsArray);
                                            bool result = grpcToolInstance.WriteDataToDataAccessService(grpcDataAccessServiceClient, dataAccessServiceWriteRequest, new IDataAccessServiceWriteResponse(), options1);
                                        }
                                        catch (Exception e)
                                        {
                                            Console.WriteLine("ERRO: {0}，{1}", e, nodeidDictionary.GetValueOrDefault(i.ToString()));
                                        }
                                        listWriteItem.Clear();
                                    }
                                    #endregion
                                }


                                TimeSpan end = new TimeSpan(DateTime.Now.Ticks);
                                DateTime nowDisplay = DateTime.Now;
                                TimeSpan dur = (start - end).Duration();

                                //logNet.WriteInfo("Thread ReadDeviceInfo read time : " + (dur.TotalMilliseconds).ToString());
                                Console.WriteLine("Thread ReadDeviceInfo read time:{0} read Duration:{1}", nowDisplay.ToString("yyyy-MM-dd HH:mm:ss:fff"), dur.TotalMilliseconds);

                                if (dur.TotalMilliseconds < 100)
                                {
                                    int sleepTime = 100 - (int)dur.TotalMilliseconds;
                                    Thread.Sleep(sleepTime);
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
                            short DM1814 = 0 ;
                            short DM1822 = 0;

                            var mc = _mc[0];


                            while(true)
                            {
                                TimeSpan start = new TimeSpan(DateTime.Now.Ticks);

                                OperateResult<short[]> ret = mc.ReadInt16("ZF200200", 8);                         
                                if (ret.IsSuccess)
                                {
                                    Array.Copy(ret.Content, ZFarray1, ret.Content.Length);
                                }
                                else
                                {
                                    logNet.WriteInfo(DateTime.Now.ToString() + ":  ZF200200 Array Read failed");
                                    Console.WriteLine(" ZF200200 Array Read failed");
                                }

                                ret = mc.ReadInt16("ZF20000", 8);
                                if (ret.IsSuccess)
                                {
                                    Array.Copy(ret.Content, ZFarray2, ret.Content.Length);
                                }
                                else
                                {
                                    logNet.WriteInfo(DateTime.Now.ToString() + ":  ZF20000 Array Read failed");
                                    Console.WriteLine(" ZF20000 Array Read failed");
                                }


                                ret = mc.ReadInt16("DM607", 72);
                                if (ret.IsSuccess)
                                {
                                    Array.Copy(ret.Content, DMarray, ret.Content.Length);
                                }
                                else
                                {
                                    logNet.WriteInfo(DateTime.Now.ToString() + ":  DMArray Read failed");
                                    Console.WriteLine(" DMArray Array Read failed");
                                }

                                OperateResult<short> ret1 = mc.ReadInt16("DM1814");
                                if (ret1.IsSuccess)
                                {
                                    DM1814 = ret1.Content;
                                }
                                else
                                {
                                    logNet.WriteInfo(DateTime.Now.ToString() + ":  DM1814 Read failed");
                                    Console.WriteLine(" DM1814 Array Read failed");
                                }

                                ret1 = mc.ReadInt16("DM1822");
                                if (ret1.IsSuccess)
                                {
                                    DM1822 = ret1.Content;
                                }
                                else
                                {
                                    logNet.WriteInfo(DateTime.Now.ToString() + ":  DM1822 Read failed");
                                    Console.WriteLine(" DM1822 Array Read failed");
                                }


                                keyenceClients.SendStationData(ZhuYeWei, DMarray, 0, grpcToolInstance, nodeidDictionary, grpcDataAccessServiceClient, options1);
                                keyenceClients.SendStationData(JingZhiWei, DMarray, DM1814, grpcToolInstance, nodeidDictionary, grpcDataAccessServiceClient, options1);
                                keyenceClients.SendStationData(FengZhuangWei, ZFarray1, ZFarray2, DMarray, DM1822, grpcToolInstance, nodeidDictionary, grpcDataAccessServiceClient, options1);


                                TimeSpan end = new TimeSpan(DateTime.Now.Ticks);
                                DateTime nowDisplay = DateTime.Now;
                                TimeSpan dur = (start - end).Duration();

                                //logNet.WriteInfo("Thread ReadStationInfo read time : " + (dur.TotalMilliseconds).ToString());
                                Console.WriteLine("Thread ReadStationInfo read time:{0} read Duration:{1}", nowDisplay.ToString("yyyy-MM-dd HH:mm:ss:fff"), dur.TotalMilliseconds);

                                if (dur.TotalMilliseconds < 100)
                                {
                                    int sleepTime = 100 - (int)dur.TotalMilliseconds;
                                    Thread.Sleep(sleepTime);
                                }

                            }

                        });

                        //读1000ms数据
                        thr[2] = new Thread(() =>
                        {
                            var mc = _mc[0];
                            var listWriteItem = new List<WriteItem>();
                            WriteItem[] writeItems = new WriteItem[] { };


                            while (true)
                            {
                                TimeSpan start = new TimeSpan(DateTime.Now.Ticks);


                                //读取并发送 生产统计、寿命管理、功能开关的数据
                                keyenceClients.ReadandSendOneSecConData(Function_Enable, mc, grpcToolInstance, nodeidDictionary, grpcDataAccessServiceClient, options1);
                                keyenceClients.ReadandSendOneSecConData(Production_Data, mc, grpcToolInstance, nodeidDictionary, grpcDataAccessServiceClient, options1);
                                keyenceClients.ReadandSendOneSecDisData(Life_Management, mc, grpcToolInstance, nodeidDictionary, grpcDataAccessServiceClient, options1);


                                #region 读取OEE数据，合并后发送
                                bool[] OEE_part1 = keyenceClients.ReadOneSecConData(OEE1, mc);
                                bool[] OEE_part2 = keyenceClients.ReadOneSecConData(OEE2, mc);
                                var IECOEENumber = 15;
                                var OEEValue = new bool[IECOEENumber];  //与IEC对应 15
                                var OEEIndex = 0;

                                //将 OEE1和OEE2 拼成一个OEE数组
                                Array.Copy(OEE_part1, 0, OEEValue, OEEIndex, OEE_part1.Length);
                                OEEIndex += OEE_part1.Length;
                                Array.Copy(OEE_part2, 0, OEEValue, OEEIndex, OEE_part2.Length);
                                
                                //Grpc发送数据给IEC
                                try
                                {
                                    listWriteItem.Add(grpcToolInstance.CreatWriteItem(nodeidDictionary["OEE"], Arp.Type.Grpc.CoreType.CtArray, OEEValue));
                                    var writeItemsArray = listWriteItem.ToArray();
                                    var dataAccessServiceWriteRequest = grpcToolInstance.ServiceWriteRequestAddDatas(writeItemsArray);
                                    bool result = grpcToolInstance.WriteDataToDataAccessService(grpcDataAccessServiceClient, dataAccessServiceWriteRequest, new IDataAccessServiceWriteResponse(), options1);
                                }
                                catch (Exception e)
                                {

                                    Console.WriteLine("ERRO: {0}", e, nodeidDictionary.GetValueOrDefault("OEE"));
                                }
                                listWriteItem.Clear();

                                #endregion

                                #region 读取报警信号数据,合并后发送

                                var ArrayIndex = 0;
                                var IECAlarmNumber = 1400;
                                var AlarmValue = new bool[IECAlarmNumber];

                                //bool[] Alarm_part1 = keyenceClients.ReadOneSecDisData(Alarm1, mc);
                                //bool[] Alarm_part2 = keyenceClients.ReadOneSecDisData(Alarm2, mc);
                                //bool[] Alarm_part3 = keyenceClients.ReadOneSecDisData(Alarm3, mc);
                                //bool[] Alarm_part4 = keyenceClients.ReadOneSecDisData(Alarm4, mc);

                                //Array.Copy(Alarm_part1, 0, AlarmValue, ArrayIndex, Alarm_part1.Length);
                                //ArrayIndex += Alarm_part1.Length;
                                //Array.Copy(Alarm_part2, 0, AlarmValue, ArrayIndex, Alarm_part2.Length);
                                //ArrayIndex += Alarm_part2.Length;
                                //Array.Copy(Alarm_part3, 0, AlarmValue, ArrayIndex, Alarm_part3.Length);
                                //ArrayIndex += Alarm_part3.Length;
                                //Array.Copy(Alarm_part4, 0, AlarmValue, ArrayIndex, Alarm_part4.Length);


                                //简化版本
                                List<bool[]> alarmGroups = new List<bool[]> {
                                                keyenceClients.ReadOneSecDisData(Alarm1, mc),
                                                keyenceClients.ReadOneSecDisData(Alarm2, mc),
                                                keyenceClients.ReadOneSecDisData(Alarm3, mc),
                                                keyenceClients.ReadOneSecDisData(Alarm4, mc) };

                                foreach (var alarmGroup in alarmGroups)
                                {
                                    Array.Copy(alarmGroup, 0, AlarmValue, ArrayIndex, alarmGroup.Length);
                                    ArrayIndex += alarmGroup.Length;
                                }

                                try
                                {
                                    listWriteItem.Add(grpcToolInstance.CreatWriteItem(nodeidDictionary["Alarm"], Arp.Type.Grpc.CoreType.CtArray, AlarmValue));
                                    var writeItemsArray = listWriteItem.ToArray();
                                    var dataAccessServiceWriteRequest = grpcToolInstance.ServiceWriteRequestAddDatas(writeItemsArray);
                                    bool result = grpcToolInstance.WriteDataToDataAccessService(grpcDataAccessServiceClient, dataAccessServiceWriteRequest, new IDataAccessServiceWriteResponse(), options1);
                                }
                                catch (Exception e)
                                {

                                    Console.WriteLine("ERRO: {0}，{1}", e, nodeidDictionary.GetValueOrDefault("Alarm"));
                                }
                                listWriteItem.Clear();

                                #endregion


                                TimeSpan end = new TimeSpan(DateTime.Now.Ticks);
                                DateTime nowDisplay = DateTime.Now;
                                TimeSpan dur = (start - end).Duration();

                                //logNet.WriteInfo("Thread ReadOneSecInfo read time : " + (dur.TotalMilliseconds).ToString());
                                Console.WriteLine("Thread ReadOneSecInfo read time:{0} read Duration:{1}", nowDisplay.ToString("yyyy-MM-dd HH:mm:ss:fff"), dur.TotalMilliseconds);

                                if (dur.TotalMilliseconds < 100)
                                {
                                    int sleepTime = 100 - (int)dur.TotalMilliseconds;
                                    Thread.Sleep(sleepTime);
                                }

                            }

                        });



                        #endregion




                        break;

                    case 100:

                        try
                        {
                            #region 开启三大数采线程

                            thr[0].Start();  // 读1000ms数据

                            thr[1].Start(); //读六大工位信息

                            thr[2].Start();  //读设备信息                                       
                            #endregion

                        }
                        catch
                        {
                            Console.WriteLine("Thread quit");
                            stepNumber = 1000;
                            break;

                        }

                        break;



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

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using HslCommunication;
using HslCommunication.Profinet.Omron;
using System.Threading;
using System.Security.Cryptography;
using System.Diagnostics.CodeAnalysis;
using System.Numerics;
using System.Runtime.CompilerServices;
using System.Collections;
using Grpc.Core;
using static Arp.Plc.Gds.Services.Grpc.IDataAccessService;
using Arp.Plc.Gds.Services.Grpc;
using Grpc.Net.Client;
using static Ph_Mc_ZhuYeJi.GrpcTool;
using System.Net.Sockets;
using System.Drawing;
using Opc.Ua;
using NPOI.SS.Formula.Functions;
using HslCommunication.LogNet;
using Microsoft.Extensions.Logging;
using static Ph_Mc_ZhuYeJi.UserStruct;
using static Ph_Mc_ZhuYeJi.Program;
using HslCommunication.Profinet.LSIS;
using static Ph_Mc_ZhuYeJi.UserStruct;
using MathNet.Numerics.LinearAlgebra.Factorization;
using HslCommunication.Profinet.Keyence;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Net.NetworkInformation;
using NPOI.Util;





namespace Ph_Mc_ZhuYeJi
{

    class KeyenceComm
    {


        #region 读取设备信息（以数组形式一起读上来，再按照序号和偏移地址写入对应的工位里）

        public void ReadDeviceInfoStruct(DeviceInfoConSturct_MC[] input, KeyenceMcNet mc, ref UDT_StationListInfo StationListInfo)
        {
            string ReadObject = input[0].varName;
            int startaddress = input[0].varOffset;

            switch (ReadObject)
            {
                case "LR2100":   //电芯记忆信号
                    {
                        ushort length = 400; //(45 - 21 + 1) * 16 = 400
                        OperateResult<bool[]> ret = mc.ReadBool(ReadObject, length);    //不确定电芯记忆信号的数据类型就是BOOL
                        if (ret.IsSuccess)
                        {
                            for (int i = 0; i < input.Length; i++)
                            {
                                var index = CalculateIndex_H(startaddress, input[i].varOffset);
                                StationListInfo.arrDataPoint[input[i].stationNumber - 1].xCellMem = ret.Content[index];
                                //Output[input[i].stationNumber - 1].Append(ret.Content[index] ? "1," : "0,");

                            }
                        }
                        else
                        {
                            logNet.WriteInfo("[MC]", ReadObject + "读取失败");
                            //Console.WriteLine(ReadObject + " Read failed");
                        }


                    }
                    break;

                case "MR20701":
                    {

                        OperateResult<bool[]> ret1 = mc.ReadBool("MR20701", 10);
                        OperateResult<bool[]> ret2 = mc.ReadBool("MR200711", 5);
                        if (ret1.IsSuccess && ret2.IsSuccess)
                        {
                            for (int i = 0; i < input.Length; i++)
                            {
                                if (input[i].varOffset - 20701 < 10)   //写死了
                                {
                                    var index = CalculateIndex_H(20701, input[i].varOffset);
                                    StationListInfo.arrDataPoint[input[i].stationNumber - 1].xCellMemClear = ret1.Content[index];

                                }
                                else
                                {
                                    var index = CalculateIndex_H(200711, input[i].varOffset);   //写死了
                                    StationListInfo.arrDataPoint[input[i].stationNumber - 1].xCellMemClear = ret2.Content[index];

                                }
                            }
                        }
                        else
                        {
                            logNet.WriteInfo("[MC]", ReadObject + "读取失败");
                            //Console.WriteLine(ReadObject + " Read failed");
                        }


                    }
                    break;

                case "MR30612":
                    {

                        Dictionary<string, bool> varResults = new Dictionary<string, bool>();
                        bool result = false;

                        for (int i = 0; i < input.Length; i++)
                        {
                            //重复的只读一次
                            ReadObject = input[i].varName;

                            if (varResults.ContainsKey(ReadObject))
                            {
                                result = varResults[ReadObject];
                            }
                            else
                            {
                                OperateResult<bool> ret = mc.ReadBool(ReadObject);
                                if (ret.IsSuccess)
                                {
                                    result = ret.Content;
                                    varResults[ReadObject] = result;
                                }
                                else
                                {
                                    logNet.WriteInfo(ReadObject + " Read failed");
                                    Console.WriteLine(ReadObject + " Read failed");
                                }
                            }

                            // 根据读取到的结果进行操作
                            StationListInfo.arrDataPoint[input[i].stationNumber - 1].xStationBusy = result;
                            //Output[input[i].stationNumber - 1].Append(result ? "1," : "0,");

                        }
                    }

                    break;


                default:
                    break;

            }
        }

        public void ReadDeivceInfoStruct(DeviceInfoDisSturct_MC[] input, KeyenceMcNet mc, ref AllDataReadfromMC allDataReadfromMC, ref UDT_StationListInfo StationListInfo)
        {

            string temp = "";
                   
            var ReadObject = input[0].varName;
            ReadObject = ReadObject.Remove(6);

            OperateResult<string> ret = mc.ReadString(ReadObject, 10);

            if (ret.IsSuccess)
            {
                //var temp = tool.ConvertIntArrayToAscii(ret.Content, 0, 9);
                temp = ret.Content;

                StationListInfo.arrDataPoint[input[0].stationNumber - 1].strCellCode = ret.Content;

                allDataReadfromMC.BarCode[0] = temp;

                //Console.WriteLine("BarCode 的 ret.Content = {0}", ret.Content);
                //logNet.WriteInfo("-------------HSL 读取到的原数据: "+  ret.Content + "------------");
                        
            }
            else
            {
                logNet.WriteInfo("[MC]", ReadObject + "读取失败");
                //Console.WriteLine(ReadObject + " Read failed");

            }

            

            


        }
        

        #endregion



        #region 拼接三大工位

        //拼接注液位 或 静置位 的数据
        public void SendStationData(StationInfoStruct_MC[] input, ref AllDataReadfromMC allDataReadfromMC, ref UDT_ProcessStationDataValue ProcessStationDataValue, short[] DMarray, short DM1814)
        {         
            var senddata = new float[input.Length];

            switch (input[0].stationName)
            {
                case "加工工位（注液位）":
                    {
                        ProcessStationDataValue.arrDataPoint[0].iDataCount = (short)input.Length;  //注液位
                        for (int i = 0; i < input.Length; i++)
                        {
                            if (input[i].varName.Length == 5)
                            {
                                var index = input[i].varOffset - 607;
                                senddata[i] = (float)(DMarray[index] / Math.Pow(10, input[i].varMagnification));
                                allDataReadfromMC.ZhuYeWeiValue[i] = senddata[i].ToString();
                                ProcessStationDataValue.arrDataPoint[0].arrDataPoint[i].StringValue = senddata[i].ToString();

                            }
                            else
                            {
                                senddata[i] = (float)(DM1814 / Math.Pow(10, input[i].varMagnification));
                                allDataReadfromMC.ZhuYeWeiValue[i] = senddata[i].ToString();
                                ProcessStationDataValue.arrDataPoint[0].arrDataPoint[i].StringValue = senddata[i].ToString();

                            }
                        }
                    }
                    break;

                case "加工工位（静置位）":
                    {

                        ProcessStationDataValue.arrDataPoint[1].iDataCount = (short)input.Length;  //静置位
                        for (int i = 0; i < input.Length; i++)
                        {
                            if (input[i].varName.Length == 5)
                            {
                                var index = input[i].varOffset - 607;
                                senddata[i] = (float)(DMarray[index] / Math.Pow(10, input[i].varMagnification));
                                allDataReadfromMC.JingZhiWeiValue[i] = senddata[i].ToString();
                                ProcessStationDataValue.arrDataPoint[1].arrDataPoint[i].StringValue = senddata[i].ToString();

                            }
                            else
                            {
                                senddata[i] = (float)(DM1814 / Math.Pow(10, input[i].varMagnification));
                                allDataReadfromMC.JingZhiWeiValue[i] = senddata[i].ToString();
                                ProcessStationDataValue.arrDataPoint[1].arrDataPoint[i].StringValue = senddata[i].ToString();
                            }
                        }


                    }
                    break;



            }                
        }

        //拼接封装工位的数据
        public void SendStationData(StationInfoStruct_MC[] input, ref AllDataReadfromMC allDataReadfromMC, ref UDT_ProcessStationDataValue ProcessStationDataValue, short[] ZFarray1, short[] ZFarray2, short[] DMarray, short DM1822)
        {
            var senddata = new float[input.Length];

            ProcessStationDataValue.arrDataPoint[2].iDataCount = (short)input.Length;  //封装位

            for (int i = 0; i < input.Length; i++)
            {
                if (input[i].varName.Length == 8)
                {

                    var index = input[i].varOffset - 200200;
                    senddata[i] = (float)(ZFarray1[index]/Math.Pow(10, input[i].varMagnification));
                    allDataReadfromMC.FengZhuangValue[i] = senddata[i].ToString();
                    ProcessStationDataValue.arrDataPoint[2].arrDataPoint[i].StringValue = senddata[i].ToString();

                }
                else if (input[i].varName.Length == 7)
                {

                    var index = input[i].varOffset - 20000;
                    senddata[i] = (float)(ZFarray2[index] / Math.Pow(10, input[i].varMagnification));
                    allDataReadfromMC.FengZhuangValue[i] = senddata[i].ToString();
                    ProcessStationDataValue.arrDataPoint[2].arrDataPoint[i].StringValue = senddata[i].ToString();

                }
                else if (input[i].varName.Length == 5)
                {

                    var index = input[i].varOffset - 607;
                    senddata[i] = (float)(DMarray[index] / Math.Pow(10, input[i].varMagnification));
                    allDataReadfromMC.FengZhuangValue[i] = senddata[i].ToString();
                    ProcessStationDataValue.arrDataPoint[2].arrDataPoint[i].StringValue = senddata[i].ToString();

                }
                else if (input[i].varName.Length == 6)
                {
                    senddata[i] = (float)(DM1822 / Math.Pow(10, input[i].varMagnification));
                    allDataReadfromMC.FengZhuangValue[i] = senddata[i].ToString();
                    ProcessStationDataValue.arrDataPoint[2].arrDataPoint[i].StringValue = senddata[i].ToString();
                }
            }
         
        }

        #endregion



        #region 读1000ms的数据

        //读功能开关、寿命管理、生产统计的数据
        public void ReadOneSecData(OneSecInfoStruct_MC[] input, KeyenceMcNet mc, ref AllDataReadfromMC allDataReadfromMC, ref DeviceDataStruct_IEC DeviceDataStruct)
        {
            switch (input[0].varName)
            {
                case "LR000":  //功能开关
                    {
                        ushort length = (ushort)input.Length;
                        OperateResult<bool[]> ret = mc.ReadBool(input[0].varName, length);

                        if (ret.IsSuccess)
                        {
                            Array.Copy(ret.Content, 0, allDataReadfromMC.FunctionEnableValue, 0, ret.Content.Length);  //写入缓存区

                            Array.Copy(ret.Content, 0, DeviceDataStruct.Value_FE, 0, ret.Content.Length);  //写入IEC结构体

                        }
                        else
                        {
                            logNet.WriteInfo("[MC]",input[0].varName + "读取失败");
                            //Console.WriteLine(input[0].varName + " Read array failed");
                        }


                    }
                    break;

                case "DM116":  // 生产统计
                    {

                        ushort length = (ushort)input.Length;
                        OperateResult<int[]> ret = mc.ReadInt32(input[0].varName, length);
 
                        if (ret.IsSuccess)
                        {
                            Array.Copy(ret.Content, 0, allDataReadfromMC.ProductionDataValue, 0, ret.Content.Length);  //写入缓存区

                            Array.Copy(ret.Content, 0, DeviceDataStruct.Value_PD, 0, ret.Content.Length);  //写入IEC结构体

                        }
                        else
                        {
                            logNet.WriteInfo("[MC]", input[0].varName + "读取失败");
                            //Console.WriteLine(input[0].varName + " Read array failed");
                        }

                    }
                    break;


                case "DM300":  //寿命管理
                    {
                        ushort length = (ushort)(input[input.Length -1 ].varOffset  - input[0].varOffset + 1);
                        OperateResult<int[]> ret = mc.ReadInt32(input[0].varName, length);


                        if (ret.IsSuccess)
                        {
                            for (int i = 0; i < input.Length; i++)
                            {
                                var index = (input[i].varOffset - input[0].varOffset) / 2;
                                //senddata[i] = ret.Content[index];
                                allDataReadfromMC.LifeManagementValue[i] = ret.Content[index];
                                DeviceDataStruct.Value_LM[i] = (uint)ret.Content[index];

                            }

                            //Array.Copy(ret.Content, 0, DeviceDataStruct.Value_LM, 0, ret.Content.Length);  //写入IEC结构体

                        }
                        else
                        {
                            logNet.WriteInfo("[MC]", input[0].varName + "读取失败");
                            //Console.WriteLine(input[0].varName + " Read array failed");
                        }
                    }
                    break ;

            }
        }

        //读取OEE的数据 
        public bool[] ReadOneSecConData(OneSecInfoStruct_MC[] input, KeyenceMcNet mc)
        {
            ushort length = (ushort)(CalculateIndex_H(input[0].varOffset, input[input.Length - 1].varOffset) + 1);
            OperateResult<bool[]> ret = mc.ReadBool(input[0].varName, length);
            if (ret.IsSuccess)
            {
                return ret.Content;
            }
            else
            {
                logNet.WriteInfo(DateTime.Now.ToString() + ":   " + input[0].varName + " Read array failed");
                Console.WriteLine(input[0].varName + " Read array failed");
                return null;
            }
        }


        //读取报警信息的数据 
        public bool[] ReadOneSecDisData(OneSecInfoStruct_MC[] input, KeyenceMcNet mc, bool isB)
        {
            bool[] senddata = new bool[input.Length];  //将数据填到里面
            ushort length = 0;
            if (!isB)
            {
                length = (ushort)(CalculateIndex_H(input[0].varOffset, input[input.Length - 1].varOffset) + 1);
                OperateResult<bool[]> ret = mc.ReadBool(input[0].varName, length);
                if (ret.IsSuccess)
                {
                    for (int i = 0; i < senddata.Length; i++)
                    {
                        var index = CalculateIndex_H(input[0].varOffset, input[i].varOffset);
                        senddata[i] = ret.Content[index];
                    }

                    return senddata;

                }
                else
                {
                    logNet.WriteInfo(DateTime.Now.ToString() + ":   " + input[0].varName + " Read array failed");
                    Console.WriteLine(input[0].varName + " Read array failed");
                    return null;
                }

            }
            else
            {
                length = (ushort)(CalculateIndex_B(input[0].varOffset, input[input.Length - 1].varOffset) + 1);
                OperateResult<bool[]> ret = mc.ReadBool(input[0].varName, length);
                if (ret.IsSuccess)
                {
                    for (int i = 0; i < senddata.Length; i++)
                    {
                        var index = CalculateIndex_B(input[0].varOffset, input[i].varOffset);
                        senddata[i] = ret.Content[index];
                    }

                    return senddata;

                }
                else
                {
                    logNet.WriteInfo(DateTime.Now.ToString() + ":   " + input[0].varName + " Read array failed");
                    Console.WriteLine(input[0].varName + " Read array failed");
                    return null;
                }
            }

        }


        #endregion


        #region 读取并发送点位名
        public void ReadPointName(OneSecInfoStruct_MC[] InputStruct, ref OneSecPointNameStruct_IEC OneSecNameStruct)
        {
            switch (InputStruct[0].varName)
            {
                case "LR000":             // 功能开关
                    {
                        OneSecNameStruct.DataCount_FE = InputStruct.Length;
                        for (int i = 0; i < InputStruct.Length; i++)
                        {
                            OneSecNameStruct.Name_FE[i].StringValue = InputStruct[i].varAnnotation;
                        }
                    }
                    break;

                case "DM116":  // 生产统计
                    {
                        OneSecNameStruct.DataCount_PD = InputStruct.Length;
                        for (int i = 0; i < InputStruct.Length; i++)
                        {
                            OneSecNameStruct.Name_PD[i].StringValue = InputStruct[i].varAnnotation;
                        }
                    }
                    break;

                case "DM300":  // 寿命管理
                    {
                        OneSecNameStruct.DataCount_LM = InputStruct.Length;
                        for (int i = 0; i < InputStruct.Length; i++)
                        {
                            OneSecNameStruct.Name_LM[i].StringValue = InputStruct[i].varAnnotation;
                        }
                    }
                    break;

                default:
                    break;
            }

        }

        //读 报警信息 和 OEE的点位名
        public void ReadPointName(List<OneSecInfoStruct_MC[]> InputStruct, string Name,int stringnumber, ref OneSecPointNameStruct_IEC OneSecNameStruct)
        {

            switch(Name)
            {
                case "Alarm":
                    {
                        var index = 0;
                        OneSecNameStruct.DataCount_ALM = stringnumber;
                        foreach (var Group in InputStruct)
                        {
                            foreach (var item in Group)
                            {
                                OneSecNameStruct.Name_ALM[index++].StringValue = item.varAnnotation;
                            }
                        }

                    }

                    break;

                case "OEE":
                    {
                        var index = 0;
                        OneSecNameStruct.DataCount_OEE = stringnumber;
                        foreach (var Group in InputStruct)
                        {
                            foreach (var item in Group)
                            {
                                OneSecNameStruct.Name_OEE[index++].StringValue = item.varAnnotation;
                            }
                        }

                    }
                    break;

                default:
                    break;
            }         
        }

        //读加工工位点位名
        public void ReadPointName(List<StationInfoStruct_MC[]> StationDataStruct, ref ProcessStationNameStruct_IEC ProcessStationNameStruct)
        {
            ProcessStationNameStruct.StationCount = (short)StationDataStruct.Count;   //写入加工工位的个数

            var i = 0;  //工位数量的索引

            foreach (var StationData in StationDataStruct)
            {
                ProcessStationNameStruct.UnitStation[i].DataCount = (short)StationData.Length;   //每个加工工位的点位数量 （不超过16个点位）
                ProcessStationNameStruct.UnitStation[i].StationNO = (short)StationData[0].StationNumber;
                ProcessStationNameStruct.UnitStation[i].StationName = StationData[0].stationName;
                var j = 0;  //每个工位里采集值的索引
                foreach (var item in StationData)
                {
                    ProcessStationNameStruct.UnitStation[i].arrKey[j].StringValue = item.varAnnotation;
                    j++;
                }
                i++;
            }
        }

        //读设备信息里的 虚拟码 和后工位码
        public void ReadandSendStaionInfo(DeviceInfoConSturct_MC[] InputStruct, GrpcTool grpcToolInstance, Dictionary<string, string> nodeidDictionary, IDataAccessServiceClient grpcDataAccessServiceClient, CallOptions options1)
        {
            var senddata = new StationInfoStruct_IEC();

            senddata.StationCount = (short)InputStruct.Length;

            for (int i = 0; i < InputStruct.Length; i++)
            {
                senddata.NextStationNO[i] = (short)InputStruct[i].nextStationNumber;
                senddata.TempCodeNO[i] = (short)InputStruct[i].pseudoCode;
            }

            var listWriteItem = new List<WriteItem>();

            try
            {
                listWriteItem.Add(grpcToolInstance.CreatWriteItem(nodeidDictionary.GetValueOrDefault("StationInfoStruct"), Arp.Type.Grpc.CoreType.CtStruct, senddata));
                var writeItemsArray = listWriteItem.ToArray();
                var dataAccessServiceWriteRequest = grpcToolInstance.ServiceWriteRequestAddDatas(writeItemsArray);
                bool result = grpcToolInstance.WriteDataToDataAccessService(grpcDataAccessServiceClient, dataAccessServiceWriteRequest, new IDataAccessServiceWriteResponse(), options1);
            }
            catch (Exception e)
            {
                logNet.WriteError("[Grpc]", " 设备信息表中的信息发送失败：" + e);
                //Console.WriteLine("ERRO: {0}", e);
            }


        }

        #endregion



        #region 十六进制的软元件 索引计算

        //根据首地址和偏移地址，得出数据在所读数据中的索引 (适用于软元件 LR MR R)
        public int CalculateIndex_H(int startaddress, int offset)
        {
            // ep: 偏移地址 205401 拆分成 2054 和 01
            var start_x = startaddress % 100; //取前半部分
            var start_y = startaddress / 100; //取后半部分

            var offset_x = offset % 100; //取前半部分
            var offset_y = offset / 100; //取后半部分

            int index = (offset_y - start_y) * 16 + (offset_x - start_x);

            return index;
        }

        ////根据首地址和偏移地址，得出数据在所读数据中的索引(适用于软元件 B)
        public int CalculateIndex_B(int startaddress, int offset)
        {
            // B 的规律为 00-0F 10-1F B2000 B2010 B2020 B 2030 B2040 B2050 B2060 B2070 B2080 B2090 B20A0 B20B0 B20C0 B20D0 B20E0 B20F0

            //将 265a 变成 26510 变成 z = 6 y = 5 x =10    (z2 - z1)* 16 *16 + (y2 - y1) * 16 + (x2 - x1)
            var gewei = startaddress % 10;
            var shiwei = (startaddress / 10) % 10;
            var start_x = shiwei * 10 + gewei;
            var start_y = (startaddress / 100) % 10;
            var start_z = (startaddress / 1000) % 10;


            gewei = offset % 10;
            shiwei = (offset / 10) % 10;
            var offset_x = shiwei * 10 + gewei;
            var offset_y = (offset / 100) % 10;
            var offset_z = (offset / 1000) % 10;


            int index = (offset_z - start_z)* 16 * 16 +(offset_y - start_y) * 16 + (offset_x - start_x);

            return index;


        }

        #endregion

        //XML标签转换 工位结构体数组的工位名是中文，为了方便XML与字典对应，需要转化为英文
        private string CN2EN(string NameCN)
        {
            string NameEN = "";

            switch (NameCN)
            {
                case "加工工位（注液位）":
                    NameEN = "zhuyewei";
                    break;

                case "加工工位（静置位）":
                    NameEN = "jingzhiwei";
                    break;

                case "加工工位（封装位）":
                    NameEN = "fengzhuangwei";
                    break;



                default:
                    break;

            }

            return NameEN;

        }





    }
}


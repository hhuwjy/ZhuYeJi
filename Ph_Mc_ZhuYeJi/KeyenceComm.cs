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





namespace Ph_Mc_ZhuYeJi
{

    class KeyenceComm
    {


        #region 读取设备信息（以数组形式一起读上来，再按照序号和偏移地址写入对应的工位里）

        //StringBuilder[] Output 的长度为工位数量的长度 Output.length = 47;
        public void ReadDeviceInfo(DeviceInfoConSturct_MC[] input, KeyenceMcNet mc, StringBuilder[] Output)
        {
            string ReadObject = input[0].varName;
            int startaddress = input[0].varOffset;

            if (ReadObject == "LR2100")
            {
                ushort length = 400; //(45 - 21 + 1) * 16 = 400
                OperateResult<bool[]> ret = mc.ReadBool(ReadObject, length);    //不确定电芯记忆信号的数据类型就是BOOL
                if (ret.IsSuccess)              
                {
                    for(int i = 0; i<input.Length; i++)
                    {
                        var index = CalculateIndex_H(startaddress, input[i].varOffset);
                        Output[input[i].stationNumber].Append(ret.Content[index] ? "1," : "0,");
                    }                                                    
                }
                else
                {
                   logNet.WriteInfo(DateTime.Now.ToString() + ":   " + ReadObject + " Read failed");
                   Console.WriteLine(ReadObject + " Read failed");
                }
            }

            else if (ReadObject == "MR20701")
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
                            Output[input[i].stationNumber].Append(ret1.Content[index] ? "1," : "0,");
                        }
                        else
                        {
                            var index = CalculateIndex_H(200711, input[i].varOffset);   //写死了
                            Output[input[i].stationNumber].Append(ret2.Content[index] ? "1," : "0,");
                        }
                    }
                }
                else
                {
                    logNet.WriteInfo(DateTime.Now.ToString() + ":   " + ReadObject + " Read failed");
                    Console.WriteLine(ReadObject + " Read failed");
                }


            }

            else if (ReadObject == "MR30612")
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
                            logNet.WriteInfo(DateTime.Now.ToString() + ":   " + ReadObject + " Read failed");
                            Console.WriteLine(ReadObject + " Read failed");
                        }
                    }

                    // 根据读取到的结果进行操作
                    Output[input[i].stationNumber].Append(result ? "1," : "0,");
                }
            }
        }

        #endregion

        #region 拼接三大工位

        //拼接注液位 或 静置位 的数据
        public void SendStationData(StationInfoStruct_MC[] input, short[] DMarray, short DM1814, GrpcTool grpcToolInstance, Dictionary<string, string> nodeidDictionary, IDataAccessServiceClient grpcDataAccessServiceClient, CallOptions options1)
        {
            var listWriteItem = new List<WriteItem>();
            WriteItem[] writeItems = new WriteItem[] { };
            var tempstring = "";
            string StationName_Now = CN2EN(input[0].stationName);

            for (int i = 0; i < input.Length; i++)
            {
                if (input[i].varName.Length == 5)
                {
                    var index = input[i].varOffset - 607;
                    tempstring += DMarray[index].ToString() + ",";
                }
                else
                {
                    tempstring += DM1814.ToString() + ",";
                }               
            }

            try
            {
                listWriteItem.Add(grpcToolInstance.CreatWriteItem(nodeidDictionary[StationName_Now], Arp.Type.Grpc.CoreType.CtString, tempstring));
                var writeItemsArray = listWriteItem.ToArray();
                var dataAccessServiceWriteRequest = grpcToolInstance.ServiceWriteRequestAddDatas(writeItemsArray);
                bool result = grpcToolInstance.WriteDataToDataAccessService(grpcDataAccessServiceClient, dataAccessServiceWriteRequest, new IDataAccessServiceWriteResponse(), options1);
            }
            catch (Exception e)
            {
                Console.WriteLine("ERRO: {0}，{1}", e, nodeidDictionary.GetValueOrDefault(StationName_Now));
            }
        }


        public void SendStationData(StationInfoStruct_MC[] input, short[] ZFarray1, short[] ZFarray2, short[] DMarray, short DM1822, GrpcTool grpcToolInstance, Dictionary<string, string> nodeidDictionary, IDataAccessServiceClient grpcDataAccessServiceClient, CallOptions options1)
        {
            var listWriteItem = new List<WriteItem>();
            WriteItem[] writeItems = new WriteItem[] { };
            var tempstring = "";
            string StationName_Now = CN2EN(input[0].stationName);

            for(int i = 0; i < input.Length; i++)
            {
                if (input[i].varName.Length == 8)
                {
                    var index = input[i].varOffset - 200200;
                    tempstring += ZFarray1[index].ToString() + ",";
                }
                else if (input[i].varName.Length == 7)
                {
                    var index = input[i].varOffset - 20000;
                    tempstring += ZFarray2[index].ToString() + ",";
                }
                else if (input[i].varName.Length == 5)
                {
                    var index = input[i].varOffset - 607;
                    tempstring += DMarray[index].ToString() + ",";
                }
                else if (input[i].varName.Length == 6)
                {
                    tempstring += DM1822.ToString() + ",";
                }
            }

            try
            {
                listWriteItem.Add(grpcToolInstance.CreatWriteItem(nodeidDictionary[StationName_Now], Arp.Type.Grpc.CoreType.CtString, tempstring));
                var writeItemsArray = listWriteItem.ToArray();
                var dataAccessServiceWriteRequest = grpcToolInstance.ServiceWriteRequestAddDatas(writeItemsArray);
                bool result = grpcToolInstance.WriteDataToDataAccessService(grpcDataAccessServiceClient, dataAccessServiceWriteRequest, new IDataAccessServiceWriteResponse(), options1);
            }
            catch (Exception e)
            {
                Console.WriteLine("ERRO: {0}，{1}", e, nodeidDictionary.GetValueOrDefault(StationName_Now));
            }


        }



        #endregion



        #region 读1000ms的数据

        //读取并发送连续地址数据 （功能开关 和 生产统计）
        public void ReadandSendOneSecConData(OneSecInfoStruct_MC[] input, KeyenceMcNet mc, GrpcTool grpcToolInstance, Dictionary<string, string> nodeidDictionary, IDataAccessServiceClient grpcDataAccessServiceClient, CallOptions options1)
        {
            var listWriteItem = new List<WriteItem>();
            WriteItem[] writeItems = new WriteItem[] { };

            if (input[0].varType == "BOOL")
            {
                ushort length = (ushort)input.Length;
                OperateResult<bool[]> ret = mc.ReadBool(input[0].varName, length);

                if (ret.IsSuccess)
                {
                    try
                    {
                        listWriteItem.Add(grpcToolInstance.CreatWriteItem(nodeidDictionary[input[0].varName], Arp.Type.Grpc.CoreType.CtArray, ret.Content));
                        var writeItemsArray = listWriteItem.ToArray();
                        var dataAccessServiceWriteRequest = grpcToolInstance.ServiceWriteRequestAddDatas(writeItemsArray);
                        bool result = grpcToolInstance.WriteDataToDataAccessService(grpcDataAccessServiceClient, dataAccessServiceWriteRequest, new IDataAccessServiceWriteResponse(), options1);
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine("ERRO: {0}，{1}", e, nodeidDictionary.GetValueOrDefault(input[0].varName));
                    }

                }
                else
                {
                    logNet.WriteInfo(DateTime.Now.ToString() + ":   " + input[0].varName + " Read array failed");
                    Console.WriteLine(input[0].varName + " Read array failed");
                }

            }

            else if (input[0].varType == "DINT")
            {
                ushort length = (ushort)input.Length;
                OperateResult<int[]> ret = mc.ReadInt32(input[0].varName, length);

                if (ret.IsSuccess)
                {
                    try
                    {
                        listWriteItem.Add(grpcToolInstance.CreatWriteItem(nodeidDictionary[input[0].varName], Arp.Type.Grpc.CoreType.CtArray, ret.Content));
                        var writeItemsArray = listWriteItem.ToArray();
                        var dataAccessServiceWriteRequest = grpcToolInstance.ServiceWriteRequestAddDatas(writeItemsArray);
                        bool result = grpcToolInstance.WriteDataToDataAccessService(grpcDataAccessServiceClient, dataAccessServiceWriteRequest, new IDataAccessServiceWriteResponse(), options1);
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine("ERRO: {0}，{1}", e, nodeidDictionary.GetValueOrDefault(input[0].varName));
                    }
                }
                else
                {
                    logNet.WriteInfo(DateTime.Now.ToString() + ":   " + input[0].varName + " Read array failed");
                    Console.WriteLine(input[0].varName + " Read array failed");
                }
            }
        }

        //读取并发送离散数据（寿命管理）
        public void ReadandSendOneSecDisData(OneSecInfoStruct_MC[] input, KeyenceMcNet mc, GrpcTool grpcToolInstance, Dictionary<string, string> nodeidDictionary, IDataAccessServiceClient grpcDataAccessServiceClient, CallOptions options1)
        {
            int[] senddata = new int[input.Length];  //将数据填到里面
            ushort length = (ushort)(CalculateIndex_H(input[0].varOffset, input[input.Length - 1].varOffset) + 1);
            var listWriteItem = new List<WriteItem>();
            WriteItem[] writeItems = new WriteItem[] { };

            OperateResult<int[]> ret = mc.ReadInt32(input[0].varName, length);
            if(ret.IsSuccess)
            {
                for(int i=0; i<senddata.Length; i++)
                {
                    var index = CalculateIndex_H(input[0].varOffset, input[i].varOffset);
                    senddata[i] = ret.Content[index];
                }

                try
                {
                    listWriteItem.Add(grpcToolInstance.CreatWriteItem(nodeidDictionary[input[0].varName], Arp.Type.Grpc.CoreType.CtArray, ret.Content));
                    var writeItemsArray = listWriteItem.ToArray();
                    var dataAccessServiceWriteRequest = grpcToolInstance.ServiceWriteRequestAddDatas(writeItemsArray);
                    bool result = grpcToolInstance.WriteDataToDataAccessService(grpcDataAccessServiceClient, dataAccessServiceWriteRequest, new IDataAccessServiceWriteResponse(), options1);
                }
                catch (Exception e)
                {
                    Console.WriteLine("ERRO: {0}，{1}", e, nodeidDictionary.GetValueOrDefault(input[0].varName));
                }
            }
            else
            {
                logNet.WriteInfo(DateTime.Now.ToString() + ":   " + input[0].varName + " Read array failed");
                Console.WriteLine(input[0].varName + " Read array failed");
            }
        }

        //读取OEE数据
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

        //读取报警数据
        public bool[] ReadOneSecDisData(OneSecInfoStruct_MC[] input, KeyenceMcNet mc)
        {
            bool[] senddata = new bool[input.Length];  //将数据填到里面
            ushort length = (ushort)(CalculateIndex_H(input[0].varOffset, input[input.Length - 1].varOffset) + 1);
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

        #endregion 



        //根据首地址和偏移地址，得出数据在所读数据中的索引 起始地址为21偏移地址为 4101   数组索引为 （41-21）*16 +（01 -00 ）
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


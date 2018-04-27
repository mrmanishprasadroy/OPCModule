using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using OpcRcw.Da;
using System.Runtime.InteropServices;


namespace TempModule
{
    public partial class Form1 : Form
    {
        internal const string SERVER_NAME       = "OPC.SimaticNET";     // local server name
        internal const string GROUP_NAME        = "CC3_MIS";            // Group name
        internal const string GROUP_NAME_RESET  = "CC3_MIS_RESET";      // Group name
        internal const int LOCALE_ID = 0x407;                           // LOCALE ID FOR ENGLISH.


        IOPCServer  pIOPCServer;        // OPC server interface pointer
        Object      pobjGroup1;         // Pointer to group object
        int         nSvrGroupID;        // server group handle for the added group
        IOPCSyncIO  pIOPCSyncIO;        // instance pointer for synchronous IO.
        // array for server handles for


        IntPtr      pItemValues = IntPtr.Zero;
        //IntPtr    pItemValues = null;
        IntPtr      pErrors     = IntPtr.Zero;
        IntPtr      pResults    = IntPtr.Zero;
        OPCITEMDEF[] ItemDefArray;
        OPCITEMDEF[] ItemDefArray_RESET;


        const int           S7_STR1_S_Send_STR1_Receive_TCM_HandshakingwithMIS_ID = 0;
        public const int    NUMBER_OF_TAGS = 1;//146;//451; // DB 1165 read only database
        public const int    vNo_OF_Tags = 1;//146;//451;   // DB 1165 read only database

        public const int    NUMBER_OF_TAGS_RESET = 1; // DB 1166 Write only database
        public const int    vNo_OF_Tags_RESET = 1;    // DB 1166 Write only database

        public string[] item = new string[NUMBER_OF_TAGS];  // Change this array to the exact No of Items to be added to the group ( Items )
        public int[] hndlSeq = new int[NUMBER_OF_TAGS];     // Change this array to the exact No of Items to be added to the group as ( Position )
        public int[] nItemSvrID = new int[NUMBER_OF_TAGS];
        public int[] nItemSvrID_RESET = new int[NUMBER_OF_TAGS_RESET];
        public string[] itemIds_RESET = new string[vNo_OF_Tags_RESET];
        string[] output;
        int i, j = 0;
        StreamWriter objWriteFile = null;
        string FilePath = "C:\\";
        string vDate1;


        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            LoadOPC();
        }
        private void LoadOPC()
        {
            try
            {
                // OPC server variables.
                Type svrComponenttyp;
                OPCITEMDEF[] ItemDefArray;

                // Initialise Group properties
                int bActive = 0;
                int dwRequestedUpdateRate = 250;
                int hClientGroup = 0;
                int dwLCID = LOCALE_ID;
                int pRevUpdateRate;
                int TimeBias = 0;
                float deadband = 0;
                int vNo_OF_Tags = 1;//341;



                ItemDefArray = new OPCITEMDEF[vNo_OF_Tags];

                // Access unmanaged COM memory
                GCHandle hTimeBias, hDeadband;

                hTimeBias = GCHandle.Alloc(TimeBias, GCHandleType.Pinned);
                hDeadband = GCHandle.Alloc(deadband, GCHandleType.Pinned);

                string[] itemIds    = new string[vNo_OF_Tags];
                string[] datas      = new string[vNo_OF_Tags];


                itemIds[S7_STR1_S_Send_STR1_Receive_TCM_HandshakingwithMIS_ID] = "S7:[STR1]S_Send.STR1.TataMis_Receive.TCM.HandshakingwithMIS";//"S7:[STR1]S_Send.STR1.TataMis_Receive.TCM.L2_Cut_Length_S1";//
  


                // 1. Connect to Simens OPC DA local server. ( IP Address = 127.0.0.1 ) ( External Ip Address as = 192.168.1.145)

                Guid iidRequiredInterface   = typeof(IOPCItemMgt).GUID;
                svrComponenttyp             = System.Type.GetTypeFromProgID(SERVER_NAME);
                pIOPCServer                 = (IOPCServer)System.Activator.CreateInstance(svrComponenttyp);

                // MessageBox.Show("Connected");


                //2. CC#3 add group function


                pIOPCServer.AddGroup(GROUP_NAME, bActive, dwRequestedUpdateRate, hClientGroup, hTimeBias.AddrOfPinnedObject(), hDeadband.AddrOfPinnedObject(), dwLCID, out nSvrGroupID, out pRevUpdateRate, ref iidRequiredInterface, out pobjGroup1);
                pIOPCSyncIO = (IOPCSyncIO)pobjGroup1;


                // nItemSvrID = new int[1];
                // nItemSvrID[0] = 1;// result.hServer;


                //3.CC#3 add item function

                for (int i = 0; i < vNo_OF_Tags; i++)
                {

                    ItemDefArray[i].szItemID = itemIds[i];
                }
                ((IOPCItemMgt)pobjGroup1).AddItems(NUMBER_OF_TAGS, ItemDefArray, out pResults, out pErrors);


                for (int i = 0; i < vNo_OF_Tags; i++)
                {

                    nItemSvrID[i] = i + 1;
                }


                //4. CC#3 Read item function

                IntPtr pItemValues = new IntPtr();
                pIOPCSyncIO.Read(OPCDATASOURCE.OPC_DS_DEVICE, NUMBER_OF_TAGS, nItemSvrID, out pItemValues, out pErrors);

                IntPtr position         = pItemValues;
                OPCITEMSTATE[] result   = new OPCITEMSTATE[vNo_OF_Tags];

                for (int i = 0; i < vNo_OF_Tags; i++)
                {

                    result[i] = (OPCITEMSTATE)Marshal.PtrToStructure(position, typeof(OPCITEMSTATE));
                    //MessageBox.Show("here" + result[i].vDataValue.ToString());
                    datas[i] = Convert.ToString(result[i].vDataValue);
                    position = (IntPtr)(position.ToInt32() + Marshal.SizeOf(typeof(OPCITEMSTATE)));/*********************************/

                }
                FuncFreeMemory();
                unloadcomobjects();

                // now as current data is available , proceed with the further functions.
                // tranfering the data to the previous variable of the TATA OPC Programe.

                output = new string[vNo_OF_Tags]; // TATA OPC Variable assigned to previous programe.

                for (int i = 0; i < vNo_OF_Tags; i++)
                {
                    output[i] = datas[i];
                }


                // Handle Strand Communication Faliure.  // Added on 12th August 2011 - 

                for (int i = 0; i < vNo_OF_Tags; i++)
                {
                    if ((output[i] == null) || (output[i] == ""))
                    {
                        output[i] = Convert.ToString(0);

                        string vDate = DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year;
                        string vDate1 = DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + "_" + DateTime.Now.Hour + "_" + DateTime.Now.Minute + "_" + DateTime.Now.Second;
                        objWriteFile = new StreamWriter(FilePath + "CC3_ACS_Errors_" + vDate + ".txt", true);
                        objWriteFile.WriteLine("Communicaiton Error:-" + itemIds[i] + "\tDate :-\t" + vDate1); // Log The Tag and associated PLC For the Communication Errors
                        objWriteFile.Close();

                    }

                }


                for (int i = 0; i < vNo_OF_Tags; i++)
                {
                    datas[i] = null;
                }


                // disconnecting Server.

                Marshal.ReleaseComObject(pIOPCServer);
                pIOPCServer = null;


            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message);

                string vDate = DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year;
                //objWriteFile = new StreamWriter(FilePath + "CC3_ACS_Errors_" + vDate + ".txt", true);
                //objWriteFile.WriteLine("Load OPC Error:-" + ex.Message + "\tDate :-\t" + vDate);
                //objWriteFile.Close();
            }
        }

        public void FuncFreeMemory()
        {
            if (pResults != IntPtr.Zero)
            {
                Marshal.FreeCoTaskMem(pResults);
                pResults = IntPtr.Zero;
            }
            if (pErrors != IntPtr.Zero)
            {
                Marshal.FreeCoTaskMem(pErrors);
                pErrors = IntPtr.Zero;
            }
        }
        public void unloadcomobjects()
        {

            Marshal.ReleaseComObject(pIOPCSyncIO);
            pIOPCSyncIO = null;
            pIOPCServer.RemoveGroup(nSvrGroupID, 0);
            Marshal.ReleaseComObject(pobjGroup1);
            pobjGroup1 = null;

        }
    }
}
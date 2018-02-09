using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Collections;
using System.Management;
using System.Management.Instrumentation;
using System.Data.OleDb;
using ZedGraph;
using System.Net.NetworkInformation;
using System.Diagnostics;
using System.Timers;
using System.IO;
using System.Security.AccessControl;
using System.Security.Principal;
using System.Threading;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        ComputersInLocalNetwork CLN;
        WMIInfoLocalHost WILH;
        WMIInfoRemoteHost WIRH;
        AuditLog AUDIL;
        //RemoteMonitor RM;
        // public ZedGraph.ZedGraphControl zedGraph1;
        // Работа с базой данных
        string filename, folderName;
        public enum ResultCode1 : uint
        {
            SuccessfullCompletion = 0,
            AccessDenied = 2,
            InsufficientPrivilege = 3,
            UnknownFailure = 4,
            PathNotFound = 9,
            InvalidParameter = 21
        }
        public Form1()
        {
            InitializeComponent();
        }


        private void comboBox1_SelectedValueChanged(object sender, EventArgs e)
        {
            checkedListBox1.Items.Clear();
            if (comboBox1.SelectedItem != null)
            {
                String str = comboBox1.SelectedItem.ToString();
                ArrayList list;
                switch (str)
                {
                    case "Локальный компьютер":
                        list = CLN.GetServerList(ComputersInLocalNetwork.SV_101_TYPES.SV_TYPE_ALL);
                        label3.Text = comboBox1.Text;
                        label3.ImageIndex = 0;
                        foreach (string name in list) checkedListBox1.Items.Add(name);
                        break;
                    case "Все компьютеры сети":
                        list = CLN.GetServerList(ComputersInLocalNetwork.SV_101_TYPES.SV_TYPE_ALL);
                        label3.Text = comboBox1.Text;
                        label3.ImageIndex = 1;
                        foreach (string name in list) checkedListBox1.Items.Add(name);
                        break;
                }
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            CLN = new ComputersInLocalNetwork();
            WILH = new WMIInfoLocalHost();
            WIRH = new WMIInfoRemoteHost();
            AUDIL = new AuditLog();
            EventLog eventLog = new EventLog("System", Environment.MachineName); eventLog.EntryWritten += new EntryWrittenEventHandler(OnEntryWritten); eventLog.EnableRaisingEvents = true;
            AUDIL.EventLogList(treeView2);
            label3.Text = comboBox1.Text;
            label3.ImageIndex = 0;
            ArrayList list;
            list = CLN.GetServerList(ComputersInLocalNetwork.SV_101_TYPES.SV_TYPE_ALL);
            checkedListBox1.Items.Clear();
            treeView1.Nodes.Clear();
            foreach (string name in list) checkedListBox1.Items.Add(name, true);
            // Получение имени локального компьютера
            string str = System.Environment.MachineName;
            //str = "Состояние локального компьютера "+str;
            // Формирование заголовка дерева информации о локальных настройках
            WILH.ComputerName = str;
            ManagementScope scope = new ManagementScope("\\\\" + str + "\\root\\cimv2", null);
            WILH.InitialWMI(treeView1, checkedListBox2, scope);
            // Выделить все настройки
            for (int index = 0; index < checkedListBox2.Items.Count; index++)
                checkedListBox2.SetItemChecked(index, true);
            WILH.fTreeView = treeView1;
            WILH.WMIAllInfo(0, scope);
            // вывод списка процессов на локальном компьютере
            WqlObjectQuery query1 = new WqlObjectQuery("SELECT * FROM Win32_Process");
            ManagementObjectSearcher find1 = new ManagementObjectSearcher(query1);
            string procname = "";
            string procowner = "";
            string procdomain = "";
            object[] parameters = new object[12];
            foreach (ManagementObject mo1 in find1.Get())
            {
                procname = mo1["Name"] as string;
                ResultCode1 result = (ResultCode1)mo1.InvokeMethod("GetOwner", parameters);
                result = (ResultCode1)mo1.InvokeMethod("GetOwner", parameters);
                if (result == ResultCode1.SuccessfullCompletion)
                {
                    procowner = parameters[0] as string;
                    procdomain = parameters[1] as string;
                }
                else
                {
                    switch (result)
                    {
                        case ResultCode1.AccessDenied:
                            procowner = "Доступ запрещен";
                            break;
                        case ResultCode1.InsufficientPrivilege:
                            procowner = "Не достаточно прав для просмотра информации";
                            break;
                        case ResultCode1.InvalidParameter:
                            procowner = "Некорректный параметр";
                            break;
                        case ResultCode1.PathNotFound:
                            procowner = "Путь не найден";
                            break;
                        case ResultCode1.UnknownFailure:
                            procowner = "Неизвестная ошибка";
                            break;
                    }
                }
            }
        }

        private void OnEntryWritten(object sender, EntryWrittenEventArgs e)
        {
            if (e.Entry.EntryType == EventLogEntryType.Error || e.Entry.EntryType == EventLogEntryType.Information)
            {
                if ((e.Entry.EventID == 4) || (e.Entry.EventID == 472) || (e.Entry.EventID == 477) || (e.Entry.EventID == 517) ||
                    (e.Entry.EventID == 624) || (e.Entry.EventID == 535) || (e.Entry.EventID == 533) || (e.Entry.EventID == 529) ||
                    (e.Entry.EventID == 632) || (e.Entry.EventID == 539) || (e.Entry.EventID == 534) || (e.Entry.EventID == 531) ||
                    (e.Entry.EventID == 636) || (e.Entry.EventID == 660) || (e.Entry.EventID == 6806) || (e.Entry.EventID == 4645) ||
                    (e.Entry.EventID == 642) || (e.Entry.EventID == 675) || (e.Entry.EventID == 681) || (e.Entry.EventID == 4728) ||
                    (e.Entry.EventID == 644) || (e.Entry.EventID == 676) || (e.Entry.EventID == 1102) || (e.Entry.EventID == 4732) ||
                    (e.Entry.EventID == 4740) || (e.Entry.EventID == 4756) || (e.Entry.EventID == 4768) || (e.Entry.EventID == 4776) ||
                    (e.Entry.EventID == 4738) || (e.Entry.EventID == 4733) || (e.Entry.EventID == 630) || (e.Entry.EventID == 200) ||
                    (e.Entry.EventID == 4124) || (e.Entry.EventID == 4226) || (e.Entry.EventID == 7901) ||
                    (e.Entry.EventID == 12294) || (e.Entry.EventID == 9095) || (e.Entry.EventID == 9097) || (e.Entry.EventID == 7023) ||
                    (e.Entry.EventID == 6183) || (e.Entry.EventID == 55) || (e.Entry.EventID == 1066) || (e.Entry.EventID == 6008) ||
                    (e.Entry.EventID == 861) || e.Entry.EventID == 7035)

                {
                    //MessageBox.Show("Внимание! В системе произошло событие   "+  e.Entry.EventID+ ".  Просмотрите журнал событий для выяснения причин.", "Система отслеживания вторжений", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    DialogResult result;
                    result = MessageBox.Show("Внимание!" + "\r\n" + "В системе произошло событие    " + e.Entry.EventID + "\r\n" + "Время     " +
                                             e.Entry.TimeGenerated.ToString() + "\r\n" + "Сообщение:  " + e.Entry.Message + "\r\n" + "\r\n" + "  Просмотрите журнал событий для выяснения причин.", "Система отслеживания вторжений", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }


        private void button3_Click(object sender, EventArgs e)
        {
            treeView1.BeginUpdate();
            treeView1.Nodes.Add("Parent");
            treeView1.Nodes[0].Nodes.Add("Child 1");
            treeView1.Nodes[0].Nodes.Add("Child 2");
            treeView1.Nodes[0].Nodes[1].Nodes.Add("Grandchild");
            treeView1.Nodes[0].Nodes[1].Nodes[0].Nodes.Add("Great Grandchild");
            treeView1.EndUpdate();
        }


        private void button10_Click(object sender, EventArgs e)
        {
            AUDIL.EventLogSee(treeView2, dataGridView3, groupBox4);

        }

        private void button11_Click(object sender, EventArgs e)
        {
            AUDIL.EventLogClear(treeView2, dataGridView3);
        }

        private void button12_Click(object sender, EventArgs e)
        {
            if (radioButton1.Checked) AUDIL.EventLogSeeFilter(treeView2, dataGridView3, textBox5, 0);
            if (radioButton2.Checked) AUDIL.EventLogSeeFilter(treeView2, dataGridView3, textBox6, 1);
            if (radioButton3.Checked) AUDIL.EventLogSeeFilter(treeView2, dataGridView3, dateTimePicker1, dateTimePicker2);
            if (radioButton4.Checked) AUDIL.EventLogSeeFilter(treeView2, dataGridView3, textBox7, 3);
        }

        private void AutoSizeRowsMode(Object sender, EventArgs es)
        {
            dataGridView3.AutoSizeRowsMode =
                DataGridViewAutoSizeRowsMode.AllCells;
        }

        private void dataGridView3_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            int colum = e.ColumnIndex;
            int row = e.RowIndex;
            string s = dataGridView3.Rows[row].Cells[colum].Value.ToString();
            MessageBox.Show(s);

        }

        private void button1_Click(object sender, EventArgs e)
        {

            AUDIL.ErrorLogSee(treeView2, dataGridView3, groupBox4);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SaveFileDialog st = new SaveFileDialog();
            string file_name = "";
            st.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
            if (st.ShowDialog() == DialogResult.OK);
            { file_name = st.FileName; }
            //string file_name = "C:\\test1.txt";

            System.IO.StreamWriter objWriter; objWriter = new System.IO.StreamWriter(file_name);

            for (int x = 0; x < dataGridView3.Columns.Count; x++)
            {
                objWriter.Write(dataGridView3.Columns[x].HeaderText);
                if (x != dataGridView3.Columns.Count - 1)
                {
                    objWriter.Write(", ");
                }

            }
            objWriter.WriteLine();

            //writing the data
            for (int x = 0; x < dataGridView3.Rows.Count - 1; x++)
            {
                for (int y = 0; y < dataGridView3.Columns.Count; y++)
                {
                    objWriter.Write(dataGridView3.Rows[x].Cells[y].Value.ToString());
                    if (y != dataGridView3.Columns.Count - 1)
                    {
                        objWriter.Write(", ");
                    }
                }
                objWriter.WriteLine();
            }
            objWriter.Close();
            MessageBox.Show("Результаты успешно сохранены", "Сохранение", MessageBoxButtons.OK, MessageBoxIcon.Information);

        }

        private void notifyIcon1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            Show();
            WindowState = FormWindowState.Normal;
            notifyIcon1.Visible = false;
        }

        private void notifyIcon1_MouseClick(object sender, MouseEventArgs e)
        {
            WindowState = FormWindowState.Minimized;
        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            if (FormWindowState.Minimized == WindowState)
            {
                notifyIcon1.Visible = true;
                Hide();
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void button9_Click(object sender, EventArgs e)
        {
            openFileDialog2.Filter = "Все файлы|*.*";
            if (openFileDialog2.ShowDialog() == DialogResult.OK)
            {
                filename = openFileDialog2.FileName;
                textBox4.Text = filename;
                FileStream stream = File.Open(filename, FileMode.Open);
                FileSecurity securityDescriptor = stream.GetAccessControl();
                AuthorizationRuleCollection rules = securityDescriptor.GetAccessRules(true, true, typeof(NTAccount));
                dataGridView2.Rows.Clear();
                foreach (AuthorizationRule rule in rules)
                {
                    FileSystemAccessRule fileRule = rule as FileSystemAccessRule;
                    dataGridView2.Rows.Add(fileRule.AccessControlType, fileRule.FileSystemRights, fileRule.IdentityReference.Value);
                }
            }
        }
    }

    //----------------------------------------------------------------
    // Класс "Список компьютеров в локальной сети"
    //-----------------------------------------------------------------
    public class ComputersInLocalNetwork
    {
        [DllImport("netapi32.dll", EntryPoint = "NetServerEnum")]
        public static extern NERR NetServerEnum(
            [MarshalAs(UnmanagedType.LPWStr)]string ServerName,
            int Level, out IntPtr BufPtr,
            int PrefMaxLen, ref int EntriesRead,
            ref int TotalEntries, SV_101_TYPES ServerType,
            [MarshalAs(UnmanagedType.LPWStr)] string Domain,
            int ResumeHandle);

        [DllImport("netapi32.dll", EntryPoint = "NetApiBufferFree")]
        public static extern NERR NetApiBufferFree(IntPtr Buffer);
        // ----------------------------------------------
        // Структура получаемой информации
        // ----------------------------------------------
        [StructLayout(LayoutKind.Sequential)]
        public struct SERVER_INFO_101
        {
            [MarshalAs(UnmanagedType.U4)]
            public uint sv101_platform_id;
            [MarshalAs(UnmanagedType.LPWStr)]
            public string sv101_name;
            [MarshalAs(UnmanagedType.U4)]
            public uint sv101_version_major;
            [MarshalAs(UnmanagedType.U4)]
            public uint sv101_version_minor;
            [MarshalAs(UnmanagedType.U4)]
            public uint sv101_type;
            [MarshalAs(UnmanagedType.LPWStr)]
            public string sv101_comment;
        }

        // ----------------------------------------------
        // список ошибок, возвращаемых NetServerEnum
        // ----------------------------------------------
        public enum NERR
        {
            NERR_Success = 0, // успех
            ERROR_ACCESS_DENIED = 5,
            ERROR_NOT_ENOUGH_MEMORY = 8,
            ERROR_BAD_NETPATH = 53,
            ERROR_NETWORK_BUSY = 54,
            ERROR_INVALID_PARAMETER = 87,
            ERROR_INVALID_LEVEL = 124,
            ERROR_MORE_DATA = 234,
            ERROR_EXTENDED_ERROR = 1208,
            ERROR_NO_NETWORK = 1222,
            ERROR_INVALID_HANDLE_STATE = 1609,
            ERROR_NO_BROWSER_SERVERS_FOUND = 6118,
        }

        // ----------------------------------------------
        /// Типы серверов
        // ----------------------------------------------
        [Flags]
        public enum SV_101_TYPES : uint
        {
            SV_TYPE_WORKSTATION = 0x00000001,
            SV_TYPE_SERVER = 0x00000002,
            SV_TYPE_SQLSERVER = 0x00000004,
            SV_TYPE_DOMAIN_CTRL = 0x00000008,
            SV_TYPE_DOMAIN_BAKCTRL = 0x00000010,
            SV_TYPE_TIME_SOURCE = 0x00000020,
            SV_TYPE_AFP = 0x00000040,
            SV_TYPE_NOVELL = 0x00000080,
            SV_TYPE_DOMAIN_MEMBER = 0x00000100,
            SV_TYPE_PRINTQ_SERVER = 0x00000200,
            SV_TYPE_DIALIN_SERVER = 0x00000400,
            SV_TYPE_XENIX_SERVER = 0x00000800,
            SV_TYPE_SERVER_UNIX = SV_TYPE_XENIX_SERVER,
            SV_TYPE_NT = 0x00001000,
            SV_TYPE_WFW = 0x00002000,
            SV_TYPE_SERVER_MFPN = 0x00004000,
            SV_TYPE_SERVER_NT = 0x00008000,
            SV_TYPE_POTENTIAL_BROWSER = 0x00010000,
            SV_TYPE_BACKUP_BROWSER = 0x00020000,
            SV_TYPE_MASTER_BROWSER = 0x00040000,
            SV_TYPE_DOMAIN_MASTER = 0x00080000,
            SV_TYPE_SERVER_OSF = 0x00100000,
            SV_TYPE_SERVER_VMS = 0x00200000,
            SV_TYPE_WINDOWS = 0x00400000,
            SV_TYPE_DFS = 0x00800000,
            SV_TYPE_CLUSTER_NT = 0x01000000,
            SV_TYPE_TERMINALSERVER = 0x02000000,
            SV_TYPE_CLUSTER_VS_NT = 0x04000000,
            SV_TYPE_DCE = 0x10000000,
            SV_TYPE_ALTERNATE_XPORT = 0x20000000,
            SV_TYPE_LOCAL_LIST_ONLY = 0x40000000,
            SV_TYPE_DOMAIN_ENUM = 0x80000000,
            SV_TYPE_ALL = 0xFFFFFFFF,
        }
        // ----------------------------------------------
        // Операционная система
        // ----------------------------------------------
        public enum PLATFORM_ID : uint
        {
            PLATFORM_ID_DOS = 300,
            PLATFORM_ID_OS2 = 400,
            PLATFORM_ID_NT = 500,
            PLATFORM_ID_OSF = 600,
            PLATFORM_ID_VMS = 700,
        }
        // ----------------------------------------------
        // Массив для списка компьютеров
        // ----------------------------------------------
        //        public ArrayList srvs;
        // ----------------------------------------------
        // получим список всех компьюетеров
        // ----------------------------------------------
        public ArrayList GetServerList(SV_101_TYPES type)
        {
            SERVER_INFO_101 si;
            IntPtr pInfo = IntPtr.Zero;
            int etriesread = 0;
            int totalentries = 0;
            ArrayList srvs = new ArrayList();
            try
            {
                NERR err = NetServerEnum(null, 101, out pInfo, -1, ref etriesread, ref totalentries, type, null, 0);
                if ((err == NERR.NERR_Success || err == NERR.ERROR_MORE_DATA) && pInfo != IntPtr.Zero)
                {
                    int ptr = pInfo.ToInt32();
                    for (int i = 0; i < etriesread; i++)
                    {
                        si = (SERVER_INFO_101)Marshal.PtrToStructure(new IntPtr(ptr), typeof(SERVER_INFO_101));
                        srvs.Add(si.sv101_name.ToString()); // добавляем имя сервера в список
                        ptr += Marshal.SizeOf(si);
                    }
                }
            }
            catch (Exception) { /* обработка ошибки ничего не делаем :(*/ }
            finally
            { // освобождаем выделенную память
                if (pInfo != IntPtr.Zero) NetApiBufferFree(pInfo);
            }
            return (srvs);
        }
    }
    //---------------------------------------------
    // Класс Локальный компьютер
    //---------------------------------------------
    public class WMIInfoLocalHost
    {
        public TreeView fTreeView;
        public string ComputerName;
        public TreeNode InfoNode;
        private int h = 0;
        public CheckedListBox fCheckedListBox;
        public System.Management.ManagementScope sc;
        private const int TextSize = 256;
        private enum ResultCode : uint
        {
            SuccessfullCompletion = 0,
            AccessDenied = 2,
            InsufficientPrivilege = 3,
            UnknownFailure = 4,
            PathNotFound = 9,
            InvalidParameter = 21
        }
        // номер ветви
        int u = 0;
        // номер листа

        public void InitialWMI(TreeView aTreeView, CheckedListBox aCheckedListBox, System.Management.ManagementScope asc)
        {
            InfoNode = new TreeNode();
            InfoNode.Text = ComputerName;
            fTreeView = aTreeView;
            fTreeView.Nodes.Add(InfoNode);
            fCheckedListBox = aCheckedListBox;
            sc = asc;
            u = 0;

        }
        //------------------------------------------------
        // Получение  информации о видеоконтроллере
        //------------------------------------------------
        public void VideoColntroller(TreeView aTreeView)
        {
            //[STAThread]  
            fTreeView = aTreeView;
            //ManagementScope sc = new ManagementScope("\\\\"+ComputerName+"\\root\\cimv2", null);
            ManagementPath ph = new ManagementPath(@"Win32_VideoController");
            ManagementClass mc = new ManagementClass(sc, ph, null);

            fTreeView.Nodes[h].Nodes.Add("Свойства видеоконтроллера");
            fTreeView.Nodes[h].Nodes[u].ImageIndex = 7;
            foreach (ManagementObject ss in mc.GetInstances())
            {

                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Наименование  :  " + ss.GetPropertyValue("Name"));
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Процессор     :  " + ss.GetPropertyValue("VideoProcessor"));
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("ВидеоПамять   :  " + ss.GetPropertyValue("AdapterRAM"));
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Дескриптор    :  " + ss.GetPropertyValue("VideoModeDescription"));
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Частота обновления экрана :" + ss.GetPropertyValue("CurrentRefreshRate"));
            }
            u++;
        }
        //---------------------------------------------
        // Получение информации о компьютере
        //---------------------------------------------
        public void ComputerSystem(TreeView aTreeView)
        {
            fTreeView = aTreeView;
            WqlObjectQuery query = new WqlObjectQuery("SELECT * FROM Win32_ComputerSystem");
            ManagementObjectSearcher find = new ManagementObjectSearcher(query);
            fTreeView.Nodes[h].Nodes.Add("Информация о компьютере");
            fTreeView.Nodes[h].Nodes[u].ImageIndex = 8;
            foreach (ManagementObject mo in find.Get())
            {
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Компьютер входит в домен : " + mo["Domain"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Фирма-изготовитель       : " + mo["Manufacturer"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Модель                   : " + mo["Model"]);
            }
            u++;
        }
        //---------------------------------------------
        // Получение информации о производителе
        //---------------------------------------------
        public void ComputerSystemProduct(TreeView aTreeView)
        {
            fTreeView = aTreeView;
            WqlObjectQuery query = new WqlObjectQuery("SELECT * FROM Win32_ComputerSystemProduct");
            ManagementObjectSearcher find = new ManagementObjectSearcher(query);
            fTreeView.Nodes[h].Nodes.Add("Информация о фирме-изготовителе");
            fTreeView.Nodes[h].Nodes[u].ImageIndex = 10;

            foreach (ManagementObject mo in find.Get())
            {
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Дескриптор               : " + mo["Description"]);
                // fTreeView.Nodes[h].Nodes[u].Nodes[l].; ; l++;
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Идентификационный номер  : " + mo["IdentifyingNumber"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Наименование продукта    : " + mo["Name"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Уникальный универсальный идентификатор : " + mo["UUID"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Торговая фирма           : " + mo["Vendor"]);
            }
            u++;
        }

        //---------------------------------------------
        // Получение информации об операционной системе
        //---------------------------------------------
        public void OperatingSystem(TreeView aTreeView)
        {
            fTreeView = aTreeView;
            WqlObjectQuery query = new WqlObjectQuery("SELECT * FROM Win32_OperatingSystem");
            ManagementObjectSearcher find = new ManagementObjectSearcher(query);
            fTreeView.Nodes[h].Nodes.Add("Информация об операционной системе");
            fTreeView.Nodes[h].Nodes[u].ImageIndex = 9;
            foreach (ManagementObject mo in find.Get())
            {
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Имя устройства, с которого производится загрузка ОС : " + mo["BootDevice"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Номер                : " + mo["BuildNumber"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Наименование         : " + mo["Caption"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Кодовая страница     : " + mo["CodeSet"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Код страны           : " + mo["CountryCode"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Версия Service Pack  : " + mo["CSDVersion"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Наименование системы : " + mo["CSName"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Временная зона       : " + mo["CurrentTimeZone"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("OS is debug build    : " + mo["Debug"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("OS is distributed across serial nodes : " + mo["Distributed"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Encryption level of transaction : " + mo["EncryptionLevel"] + " bits");
                fTreeView.Nodes[h].Nodes[u].Nodes.Add(">>Priority increase for foreground app : " + GetForeground(mo));
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Доступная физическая память : " + mo["FreePhysicalMemory"] + " килобайт");
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Доступная виртуальная память : " + mo["FreeVirtualMemory"] + " килобайт");
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Количество свободных cтраниц в страничном файле : " + mo["FreeSpaceInPagingFiles"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Дата инсталляции ОС : " + ManagementDateTimeConverter.ToDateTime(mo["InstallDate"].ToString()));
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Тип оптимизации памяти : " + (Convert.ToInt16(mo["LargeSystemCache"]) == 0 ? "for applications" : "for system perfomance"));
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Время с последней загрузки : " + mo["LastBootUpTime"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Местные дата и время : " + ManagementDateTimeConverter.ToDateTime(mo["LocalDateTime"].ToString()));
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Идентификатор языка : " + mo["Locale"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Макс. доп. количество процессов : " + mo["MaxNumberOfProcesses"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Макс. память,выделяемая процессу : " + mo["MaxProcessmemorySize"] + " Кбайт");
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Колич. процессов в данный момент : " + mo["NumberOfProcesses"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Колич.пользователей в данный момент : " + mo["NumberOfUsers"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("OS laguage version : " + mo["OSLanguage"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("OS product suite version : " + GetSuite(mo));
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Number of Windows Plus : " + mo["PlusProductID"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Version of Windows Plus : " + mo["PlusVersionNumber"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Type of installed OS : " + GetProductType(mo));
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Registered user of OS : " + mo["RegisteredUser"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Серийный номер ОС : " + mo["SerialNumber"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("ServicePack major version : " + mo["ServicePackMajorVersion"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Total number to store in paging files : " + mo["SizeStoredInPagingFiles"] + "Кбайт");
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Статус : " + mo["Status"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("OS suite : " + GetOSSuite(mo));
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Загрузочный диск : " + mo["SystemDevice"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Системный каталог : " + mo["SystemDirectory"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Каталог с Windows : " + mo["WindowsDirectory"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Общая виртуальная память : " + mo["TotalVirtualMemorySize"] + "Кбайт");
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Общая физическая память : " + mo["TotalVisibleMemorySize"] + "Кбайт");
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Версия ОС : " + mo["Version"]);
            }
            u++;
        }
        private static string GetForeground(ManagementObject mo)
        {
            int i = Convert.ToInt16(mo["ForegroundApplicationBoost"]);
            switch (i)
            {
                case 0: return "None";
                case 1: return "Minimum";
                case 2: return "maximum (default value)";
            }
            return "Boost not defined";
        }
        private static string GetSuite(ManagementObject mo)
        {
            uint i = Convert.ToUInt32(mo["OSProductSuite"]);
            switch (i)
            {
                case 1: return "Small Business";
                case 2: return "Enterprise";
                case 4: return "BackOffice";
                case 8: return "Communication Server";
                case 16: return "Terminal Server";
                case 32: return "Small Business (Restricted)";
                case 64: return "Embedded NT";
                case 128: return "Data Center";
            }
            return "Os suite not defined";
        }
        //---------------------------------------------
        // Тип ОС
        //---------------------------------------------
        private static string GetOSType(ManagementObject mo)
        {
            uint i = Convert.ToUInt16(mo["OSType"]);
            switch (i)
            {
                case 16: return "WIN95";
                case 17: return "WIN98";
                case 18: return "WINNT";
                case 19: return "WINCE";
            }
            return "Тип ОС не определен";
        }
        private static string GetProductType(ManagementObject mo)
        {
            uint i = Convert.ToUInt32(mo["ProductType"]);
            switch (i)
            {
                case 1: return "Рабочая станция";
                case 2: return "Контроллер домена";
                case 3: return "Сервер";
            }
            return "Тип не определен";
        }
        private static string GetOSSuite(ManagementObject mo)
        {
            uint i = Convert.ToUInt32(mo["SuiteMask"]);
            string suite = "";
            if ((i & 1) == 1) suite += "Small Business";
            if ((i & 2) == 2)
            {
                if (suite.Length > 0) suite += ", "; suite += " Enterprise ";
            }
            if ((i & 4) == 4)
            {
                if (suite.Length > 0) suite += ", "; suite += " Back Office ";
            }
            if ((i & 8) == 8)
            {
                if (suite.Length > 0) suite += ", "; suite += " Communications ";
            }
            if ((i & 16) == 16)
            {
                if (suite.Length > 0) suite += ", "; suite += " Terminal ";
            }
            if ((i & 32) == 32)
            {
                if (suite.Length > 0) suite += ", "; suite += " Small Business Restricted ";
            }
            if ((i & 64) == 64)
            {
                if (suite.Length > 0) suite += ", "; suite += " Embedded NT ";
            }
            if ((i & 128) == 128)
            {
                if (suite.Length > 0) suite += ", "; suite += " Data Center";
            }
            if ((i & 256) == 256)
            {
                if (suite.Length > 0) suite += ", "; suite += " Single User ";
            }
            if ((i & 512) == 512)
            {
                if (suite.Length > 0) suite += ", "; suite += " Personal ";
            }
            if ((i & 1024) == 1024)
            {
                if (suite.Length > 0) suite += ", "; suite += " Blade ";
            }
            return suite;
        }

        //---------------------------------------------
        // Получение информации о компьютере
        //---------------------------------------------
        public void ComputerSystem_1(TreeView aTreeView)
        {
            fTreeView = aTreeView;
            string[] Roles =
            {
                "Standalone Workstation",    //0
                "Member Workstation",        //1
                "Standalone Server",         //2
                "Member Server",             //3
                "Backup Domain Controller",  //4
                "Primary Domain Controller"  //5
            };
            WqlObjectQuery query = new WqlObjectQuery("SELECT * FROM Win32_ComputerSystem");
            ManagementObjectSearcher find = new ManagementObjectSearcher(query);
            fTreeView.Nodes[h].Nodes.Add("Общая информация  о компьютере");
            fTreeView.Nodes[h].Nodes[u].ImageIndex = 11;
            foreach (ManagementObject mo in find.Get())
            {
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("AdminPasswordStatus: " + mo["AdminPasswordStatus"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("AutomaticResetBootOption:" + mo["AutomaticResetBootOption"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("AutomaticResetCapability: " + mo["AutomaticResetCapability"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("BootOptionOnLimit: " + mo["BootOptionOnLimit"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("BootOptionOnWatchDog: " + mo["BootOptionOnWatchDog"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("BootROMSupported: " + mo["BootROMSupported"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("BootupState: " + mo["BootupState"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Caption: " + mo["Caption"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("ChassisBootupState: " + mo["ChassisBootupState"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("CreationClassName: " + mo["CreationClassName"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("CurrentTimeZone: " + mo["CurrentTimeZone"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("DaylightInEffect: " + mo["DaylightInEffect"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Description: " + mo["Description"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Domain: " + mo["Domain"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("DomainRole: " + mo["DomainRole"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("EnableDaylightSavingsTime: " + mo["EnableDaylightSavingsTime"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("FrontPanelResetStatus: " + mo["FrontPanelResetStatus"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("InfraredSupported: " + Roles[Convert.ToInt32(mo["InfraredSupported"])]);
                if (mo["InitialLoadInfo"] == null)
                    fTreeView.Nodes[h].Nodes[u].Nodes.Add("InitialLoadInfo: " + mo["InitialLoadInfo"]);
                else
                {
                    String[] arrInitialLoadInfo = (String[])(mo["InitialLoadInfo"]);
                    foreach (String arrValue in arrInitialLoadInfo)
                    {
                        fTreeView.Nodes[h].Nodes[u].Nodes.Add("InitialLoadInfo: " + arrValue);
                    }
                }
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("InstallDate: " + mo["InstallDate"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("KeyboardPasswordStatus: " + mo["KeyboardPasswordStatus"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("LastLoadInfo: " + mo["LastLoadInfo"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Manufacturer: " + mo["Manufacturer"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Model: " + mo["Model"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Name: " + mo["Name"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("NameFormat: " + mo["NameFormat"]);
            }
            u++;
        }

        //---------------------------------------------
        // Принцип обслуживания
        //---------------------------------------------
        private static string GetServicePhilosophy(ManagementObject mo)
        {
            int i = Convert.ToInt16(mo["ServicePhilosophy"]);
            switch (i)
            {
                case 0: return "Unknown";
                case 1: return "Other";
                case 2: return "Service From Top";
                case 3: return "Service From Front";
                case 4: return "Service From Back";
                case 5: return "Service From Side";
                case 6: return "Sliding Trays";
                case 7: return "Removable Sides";
                case 8: return "Moveable";
            }
            return "Service philosophy not defined";
        }

        //---------------------------------------------
        // Физические параметры компьютера
        //---------------------------------------------
        public void SystemEnclosure(TreeView aTreeView)
        {
            fTreeView = aTreeView;
            WqlObjectQuery query = new WqlObjectQuery("SELECT * FROM Win32_SystemEnclosure");
            ManagementObjectSearcher find = new ManagementObjectSearcher(query);
            fTreeView.Nodes[h].Nodes.Add("Физические параметры компьютера");
            fTreeView.Nodes[h].Nodes[u].ImageIndex = 12;
            foreach (ManagementObject mo in find.Get())
            {
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("AudibleAlarm: " + mo["AudibleAlarm"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("BreachDescription: " + mo["BreachDescription"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("CableManagementStrategy:" + mo["CableManagementStrategy"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Caption:" + mo["Caption"]);

                fTreeView.Nodes[h].Nodes[u].Nodes.Add("CreationClassName: " + mo["CreationClassName"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("CurrentRequiredOrProduced: " + mo["CurrentRequiredOrProduced"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Depth: " + mo["Depth"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Description: " + mo["Description"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("HeatGeneration: " + mo["HeatGeneration"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Height: " + mo["Height"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("HotSwappable: " + mo["HotSwappable"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("InstallDate: " + mo["InstallDate"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("LockPresent: " + mo["LockPresent"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Manufacturer: " + mo["Manufacturer"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Model: " + mo["Model"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Name: " + mo["Name"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("NumberOfPowerCords: " + mo["NumberOfPowerCords"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("OtherIdentifyingInfo: " + mo["OtherIdentifyingInfo"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("PartNumber: " + mo["PartNumber"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("PoweredOn: " + mo["PoweredOn"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Removable: " + mo["Removable"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Replaceable: " + mo["Replaceable"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("SecurityBreach: " + mo["SecurityBreach"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("SecurityStatus: " + mo["SecurityStatus"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("SerialNumber: " + mo["SerialNumber"]);
                if (mo["ServiceDescriptions"] == null)
                    fTreeView.Nodes[h].Nodes[u].Nodes.Add("ServiceDescriptions: " + mo["ServiceDescriptions"]);
                else
                {
                    String[] arrServiceDescriptions = (String[])(mo["ServiceDescriptions"]);
                    foreach (String arrValue in arrServiceDescriptions)
                    {
                        fTreeView.Nodes[h].Nodes[u].Nodes.Add("ServiceDescriptions: " + arrValue);
                    }
                }
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("ServicePhilosophy: " + GetServicePhilosophy(mo));
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("SKU: " + mo["SKU"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("SMBIOSAssetTag: " + mo["SMBIOSAssetTag"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Status: " + mo["Status"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Tag: " + mo["Tag"]);
                if (mo["TypeDescriptions"] == null)
                    fTreeView.Nodes[h].Nodes[u].Nodes.Add("TypeDescriptions: " + mo["TypeDescriptions"]);
                else
                {
                    String[] arrTypeDescriptions = (String[])(mo["TypeDescriptions"]);
                    foreach (String arrValue in arrTypeDescriptions)
                    {
                        fTreeView.Nodes[h].Nodes[u].Nodes.Add("TypeDescriptions: ", arrValue);
                    }
                }
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Version: " + mo["Version"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("VisibleAlarm: " + mo["VisibleAlarm"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Weight: " + mo["Weight"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Width: " + mo["Width"]);
            }
            u++;
        }
        //---------------------------------------------
        // Определение архитектуры процессора
        //---------------------------------------------
        private static string GetArchitecture(ManagementBaseObject mo)
        {
            int i = Convert.ToInt32(mo["Architecture"]);
            switch (i)
            {
                case 0: return "x86";
                case 1: return "MIPS";
                case 2: return "Alpha";
                case 3: return "PowerPC";
                case 4: return "ia64";
            }
            return "Undefined";
        }
        //---------------------------------------------
        // Определение состояния ЦПУ
        //---------------------------------------------
        private static string GetCPUStatus(ManagementBaseObject mo)
        {
            int i = Convert.ToInt32(mo["CpuStatus"]);
            switch (i)
            {
                case 0: return "Unknown";
                case 1: return "CPU Enabled";
                case 2: return "CPU Disabled by User via BIOS Setup";
                case 3: return "CPU Disabled by BIOS (POST Error)";
                case 4: return "CPU is Idle";
                case 5: return "This value is reserved";
                case 6: return "This value is reserved";
                case 7: return "Other";
            }
            return "Undefined";
        }

        //---------------------------------------------
        // Определение типа процессора
        //---------------------------------------------
        private static string GetProcessorType(ManagementBaseObject mo)
        {
            int i = Convert.ToInt32(mo["ProcessorType"]);
            switch (i)
            {
                case 1: return "Other";
                case 2: return "Unknown";
                case 3: return "Central processor";
                case 4: return "Math Processor";
                case 5: return "DSP Processor";
                case 6: return "Video Processor";
            }
            return "Undefined type";
        }
        //---------------------------------------------
        // Определение состояния процессора
        //---------------------------------------------
        private static string GetStatusInfo(ManagementBaseObject mo)
        {
            int i = Convert.ToInt32(mo["StatusInfo"]);
            switch (i)
            {
                case 1: return "Other";
                case 2: return "Unknown";
                case 3: return "Enabled";
                case 4: return "Disabled";
                case 5: return "Not applicable";
            }
            return "StatusInfo not defined";
        }
        //---------------------------------------------
        // Получение информации о процессорах
        //---------------------------------------------
        public void Processor(TreeView aTreeView)
        {
            fTreeView = aTreeView;
            WqlObjectQuery query = new WqlObjectQuery("SELECT * FROM Win32_Processor");
            ManagementObjectSearcher find = new ManagementObjectSearcher(query);
            fTreeView.Nodes[h].Nodes.Add("Информация о процессорах");
            fTreeView.Nodes[h].Nodes[u].ImageIndex = 14;
            foreach (ManagementObject mo in find.Get())
            {
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("AddressWidth: " + mo["AddressWidth"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Architecture: " + GetArchitecture(mo));
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Availability: " + mo["Availability"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Caption: " + mo["Caption"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("ConfigManagerErrorCode: " + mo["ConfigManagerErrorCode"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("ConfigManagerUserConfig: " + mo["ConfigManagerUserConfig"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("CpuStatus: " + GetCPUStatus(mo));
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("CreationClassName: " + mo["CreationClassName"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("CurrentClockSpeed: " + mo["CurrentClockSpeed"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("CurrentVoltage: " + mo["CurrentVoltage"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("DataWidth: " + mo["DataWidth"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Description: " + mo["Description"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("DeviceID: " + mo["DeviceID"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("ErrorCleared: " + mo["ErrorCleared"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("ErrorDescription: " + mo["ErrorDescription"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("ExtClock: " + mo["ExtClock"]);
                // // fTreeView.Nodes[h].Nodes[u].Nodes.Add("Family: " + GetFamily(mo));
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("InstallDate: " + mo["InstallDate"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("L2CacheSize: " + mo["L2CacheSize"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("L2CacheSpeed: " + mo["L2CacheSpeed"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("LastErrorCode: " + mo["LastErrorCode"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Level: " + mo["Level"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("LoadPercentage: " + mo["LoadPercentage"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Manufacturer: " + mo["Manufacturer"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("MaxClockSpeed: " + mo["MaxClockSpeed"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Name: " + mo["Name"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("OtherFamilyDescription: " + mo["OtherFamilyDescription"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("PNPDeviceID: " + mo["PNPDeviceID"]);
                if (mo["PowerManagementCapabilities"] == null)
                    fTreeView.Nodes[h].Nodes[u].Nodes.Add("PowerManagementCapabilities: " + mo["PowerManagementCapabilities"]);
                else
                {
                    UInt16[] arrPowerManagementCapabilities = (UInt16[])(mo["PowerManagementCapabilities"]);
                    foreach (UInt16 arrValue in arrPowerManagementCapabilities)
                    {
                        fTreeView.Nodes[h].Nodes[u].Nodes.Add("PowerManagementCapabilities: " + arrValue);
                    }
                }
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("PowerManagementSupported: " + mo["PowerManagementSupported"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("ProcessorId: " + mo["ProcessorId"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("ProcessorType: " + GetProcessorType(mo));
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Revision: " + mo["Revision"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Role: " + mo["Role"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("SocketDesignation: " + mo["SocketDesignation"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Status: " + mo["Status"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("StatusInfo: " + GetStatusInfo(mo));
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Stepping: " + mo["Stepping"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("SystemCreationClassName: " + mo["SystemCreationClassName"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("SystemName: " + mo["SystemName"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("UniqueId: " + mo["UniqueId"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("UpgradeMethod: " + mo["UpgradeMethod"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Version: " + mo["Version"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("VoltageCaps: " + mo["VoltageCaps"]);
            }
            u++;
        }
        //---------------------------------------------
        // Получение списка общих файлов и каталогов
        //---------------------------------------------
        public void ShareFilesAndDirect(TreeView aTreeView)
        {
            fTreeView = aTreeView;
            WqlObjectQuery query = new WqlObjectQuery("SELECT * FROM Win32_Share");
            ManagementObjectSearcher find = new ManagementObjectSearcher(query);
            fTreeView.Nodes[h].Nodes.Add("Список общих файлов и каталогов");
            fTreeView.Nodes[h].Nodes[u].ImageIndex = 5;
            foreach (ManagementObject mo in find.Get())
            {
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("List of shares=" + mo["Name"]);
            }
            u++;
        }
        //---------------------------------------------
        // Список подключенных сетевых ресурсов
        //---------------------------------------------
        public void NetworkConnection(TreeView aTreeView)
        {
            fTreeView = aTreeView;
            fTreeView.Nodes[h].Nodes.Add("Список подключенных сетевых ресурсов");
            fTreeView.Nodes[h].Nodes[u].ImageIndex = 4;
            WqlObjectQuery query = new WqlObjectQuery("SELECT * FROM Win32_NetworkConnection");
            ManagementObjectSearcher find = new ManagementObjectSearcher(query);
            foreach (ManagementObject mo in find.Get())
            {
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("AccessMask: " + mo["AccessMask"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Caption: " + mo["Caption"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Comment: " + mo["Comment"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("ConnectionState: " + mo["ConnectionState"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("ConnectionType: " + mo["ConnectionType"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Description: " + mo["Description"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("DisplayType: " + mo["DisplayType"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("InstallDate: " + mo["InstallDate"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("LocalName: " + mo["LocalName"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Name: " + mo["Name"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Persistent: " + mo["Persistent"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("ProviderName: " + mo["ProviderName"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("RemoteName: " + mo["RemoteName"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("RemotePath: " + mo["RemotePath"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("ResourceType: " + mo["ResourceType"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Status: " + mo["Status"]);
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("UserName: " + mo["UserName"]);
            }
            u++;
        }

        //---------------------------------------------
        // Получение информации о загрузчике системы
        //---------------------------------------------
        public void BootCofiguration(TreeView aTreeView)
        {
            fTreeView = aTreeView;
            fTreeView.Nodes[h].Nodes.Add("Информация о загрузчике системы");
            fTreeView.Nodes[h].Nodes[u].ImageIndex = 3;
            WqlObjectQuery query = new WqlObjectQuery("SELECT * FROM Win32_BootConfiguration");
            ManagementObjectSearcher find = new ManagementObjectSearcher(query);
            foreach (ManagementObject mo in find.Get())
            {
                if (mo["BootDirectory"] != null) fTreeView.Nodes[h].Nodes[u].Nodes.Add("Загрузка из каталога: " + mo["BootDirectory"]);//BootDirectory
                if (mo["Caption"] != null) fTreeView.Nodes[h].Nodes[u].Nodes.Add("Название загрузочного раздела : " + mo["Caption"]);                  //Caption
                if (mo["ConfigurationPath"] != null) fTreeView.Nodes[h].Nodes[u].Nodes.Add("Полный путь к каталогу, из которого производится загрузка ОС: " + mo["ConfigurationPath"]); //ConfigurationPath
                if (mo["Description"] != null) fTreeView.Nodes[h].Nodes[u].Nodes.Add("Описание: " + mo["Description"]); //Description
                if (mo["LastDrive"] != null) fTreeView.Nodes[h].Nodes[u].Nodes.Add("Допустимое имя последнего диска: " + mo["LastDrive"]);  //LastDrive
                if (mo["Name"] != null) fTreeView.Nodes[h].Nodes[u].Nodes.Add("Название: " + mo["Name"]); //Name
                if (mo["ScratchDirectory"] != null) fTreeView.Nodes[h].Nodes[u].Nodes.Add("Каталог с файлом сбойных секторов : " + mo["ScratchDirectory"]); //ScratchDirectory
                if (mo["SettingID"] != null) fTreeView.Nodes[h].Nodes[u].Nodes.Add("Идентификатор: " + mo["SettingID"]);//SettingID
                if (mo["TempDirectory"] != null) fTreeView.Nodes[h].Nodes[u].Nodes.Add("Временный каталог: " + mo["TempDirectory"]); //TempDirectory
            }
            u++;
        }
        //---------------------------------------------
        // Список процессов
        //---------------------------------------------
        public void Process(TreeView aTreeView)
        {
            fTreeView = aTreeView;
            fTreeView.Nodes[h].Nodes.Add("Список процессов");
            fTreeView.Nodes[h].Nodes[u].ImageIndex = 17;

            WqlObjectQuery query = new WqlObjectQuery("SELECT * FROM Win32_Process");
            ManagementObjectSearcher find = new ManagementObjectSearcher(query);
            int w = 0;
            foreach (ManagementObject mo in find.Get())
            {
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Процесс:  " + mo["Name"] as string);
                object[] parameters = new object[12];
                ResultCode result = (ResultCode)mo.InvokeMethod("GetOwner", parameters);
                if (result == ResultCode.SuccessfullCompletion)
                {
                    fTreeView.Nodes[h].Nodes[u].Nodes[w].Nodes.Add("Пользователь: " + parameters[0]);
                    fTreeView.Nodes[h].Nodes[u].Nodes[w].Nodes.Add("Домен: " + parameters[1]);
                }
                else
                {
                    string s = "";
                    switch (result)
                    {
                        case ResultCode.AccessDenied:
                            s = "Доступ запрещен";
                            break;
                        case ResultCode.InsufficientPrivilege:
                            s = "Не достаточно прав для просмотра информации";
                            break;
                        case ResultCode.InvalidParameter:
                            s = "Некорректный параметр";
                            break;
                        case ResultCode.PathNotFound:
                            s = "Путь не найден";
                            break;
                        case ResultCode.UnknownFailure:
                            s = "Неизвестная ошибка";
                            break;
                    }
                    fTreeView.Nodes[h].Nodes[u].Nodes[w].Nodes.Add("Ошибка: " + s);
                }
                w++;
            }
            u++;
        }


        //---------------------------------------------
        // Список запущенных и остановленных сервисов
        //---------------------------------------------
        public void Process1(TreeView aTreeView)
        {
            fTreeView = aTreeView;
            fTreeView.Nodes[h].Nodes.Add("Список запущенных и остановленных сервисов");
            fTreeView.Nodes[h].Nodes[u].ImageIndex = 13;
            WqlObjectQuery query = new WqlObjectQuery("SELECT * FROM Win32_Service WHERE state='running'");
            ManagementObjectSearcher find = new ManagementObjectSearcher(query);
            int w = 0;

            foreach (ManagementObject mo in find.Get())
            {
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Сервис:  " + mo["PathName"]);
                if (mo["AcceptPause"] != null) fTreeView.Nodes[h].Nodes[u].Nodes[w].Nodes.Add("Пауза разрешена : " + mo["AcceptPause"]);
                if (mo["AcceptStop"] != null) fTreeView.Nodes[h].Nodes[u].Nodes[w].Nodes.Add("Остановка разрешена: " + mo["AcceptStop"]);
                if (mo["Caption"] != null) fTreeView.Nodes[h].Nodes[u].Nodes[w].Nodes.Add("Заголовок : " + mo["Caption"]);
                if (mo["CheckPoint"] != null) fTreeView.Nodes[h].Nodes[u].Nodes[w].Nodes.Add("CheckPoint: " + mo["CheckPoint"]);
                if (mo["CreationClassName"] != null) fTreeView.Nodes[h].Nodes[u].Nodes[w].Nodes.Add("Имя класса при создании: " + mo["CreationClassName"]);
                if (mo["Description"] != null) fTreeView.Nodes[h].Nodes[u].Nodes[w].Nodes.Add("Описание: " + mo["Description"]);
                if (mo["DesktopInteract"] != null) fTreeView.Nodes[h].Nodes[u].Nodes[w].Nodes.Add("DesktopInteract: " + mo["DesktopInteract"]);
                if (mo["DisplayName"] != null) fTreeView.Nodes[h].Nodes[u].Nodes[w].Nodes.Add("Дисплей: " + mo["DisplayName"]);
                if (mo["ErrorControl"] != null) fTreeView.Nodes[h].Nodes[u].Nodes[w].Nodes.Add("ErrorControl: " + mo["ErrorControl"]);
                if (mo["ExitCode"] != null) fTreeView.Nodes[h].Nodes[u].Nodes[w].Nodes.Add("Код завершения : " + mo["ExitCode"]);
                if (mo["InstallDate"] != null) fTreeView.Nodes[h].Nodes[u].Nodes[w].Nodes.Add("Дата инсталляции: " + mo["InstallDate"]);
                if (mo["Name"] != null) fTreeView.Nodes[h].Nodes[u].Nodes[w].Nodes.Add("Название: " + mo["Name"]);
                if (mo["ProcessId"] != null) fTreeView.Nodes[h].Nodes[u].Nodes[w].Nodes.Add("Идентификатор процесса: " + mo["ProcessId"]);
                if (mo["ServiceSpecificExitCode"] != null) fTreeView.Nodes[h].Nodes[u].Nodes[w].Nodes.Add("Специальный код завершения: " + mo["ServiceSpecificExitCode"]);
                if (mo["ServiceType"] != null) fTreeView.Nodes[h].Nodes[u].Nodes[w].Nodes.Add("Тип сервиса: " + mo["ServiceType"]);
                if (mo["Started"] != null) fTreeView.Nodes[h].Nodes[u].Nodes[w].Nodes.Add("Запущен:  " + mo["Started"]);
                if (mo["StartMode"] != null) fTreeView.Nodes[h].Nodes[u].Nodes[w].Nodes.Add("Режим запуска: " + mo["StartMode"]);
                if (mo["StartName"] != null) fTreeView.Nodes[h].Nodes[u].Nodes[w].Nodes.Add("Состояние при запуске : " + mo["StartName"]);
                if (mo["State"] != null) fTreeView.Nodes[h].Nodes[u].Nodes[w].Nodes.Add("Состояние : " + mo["State"]);
                if (mo["Status"] != null) fTreeView.Nodes[h].Nodes[u].Nodes[w].Nodes.Add("Статус : " + mo["Status"]);
                if (mo["SystemCreationClassName"] != null) fTreeView.Nodes[h].Nodes[u].Nodes[w].Nodes.Add("Для какой ОС создан : " + mo["SystemCreationClassName"]);
                if (mo["SystemName"] != null) fTreeView.Nodes[h].Nodes[u].Nodes[w].Nodes.Add("Название ОС: " + mo["SystemName"]);
                if (mo["TagId"] != null) fTreeView.Nodes[h].Nodes[u].Nodes[w].Nodes.Add("Идентификатор тега: " + mo["TagId"]);
                if (mo["WaitHint"] != null) fTreeView.Nodes[h].Nodes[u].Nodes[w].Nodes.Add("WaitHint: " + mo["WaitHint"]);
                w++;
            }
            u++;
        }
        //---------------------------------------------
        // Информация о переменных окружения
        //---------------------------------------------
        public void Environment(TreeView aTreeView)
        {
            fTreeView = aTreeView;
            fTreeView.Nodes[h].Nodes.Add("Информация о переменных окружения");
            fTreeView.Nodes[h].Nodes[u].ImageIndex = 16;
            WqlObjectQuery query = new WqlObjectQuery("SELECT * FROM Win32_Environment");
            ManagementObjectSearcher find = new ManagementObjectSearcher(query);
            int w = 0;
            foreach (ManagementObject mo in find.Get())
            {

                if (mo["Name"] != null) fTreeView.Nodes[h].Nodes[u].Nodes.Add("Переменная окружения: " + mo["Name"]);
                if (mo["Description"] != null) fTreeView.Nodes[h].Nodes[u].Nodes[w].Nodes.Add("Описание: " + mo["Description"]);
                if (mo["UserName"] != null) fTreeView.Nodes[h].Nodes[u].Nodes[w].Nodes.Add("Имя пользователя: " + mo["UserName"]);
                if (mo["VariableValue"] != null) fTreeView.Nodes[h].Nodes[u].Nodes[w].Nodes.Add("Значение переменной: " + mo["VariableValue"]);
                w++;
            }
            u++;
        }
        //---------------------------------------------
        // Информация о разделах диска
        //---------------------------------------------
        public void DiskPartition(TreeView aTreeView)
        {
            fTreeView = aTreeView;
            fTreeView.Nodes[h].Nodes.Add("Разделы диска");
            fTreeView.Nodes[h].Nodes[u].ImageIndex = 18;
            WqlObjectQuery query = new WqlObjectQuery("SELECT * FROM Win32_DiskPartition");
            ManagementObjectSearcher find = new ManagementObjectSearcher(query);
            int w = 0;
            int v = 0;

            foreach (ManagementObject mo in find.Get())
            {
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Наименование диска: " + mo["Name"]);
                if (mo["Access"] != null) fTreeView.Nodes[h].Nodes[u].Nodes[w].Nodes.Add("Доступ: " + mo["Access"]);
                if (mo["Availability"] != null) fTreeView.Nodes[h].Nodes[u].Nodes[w].Nodes.Add("Доступ: " + mo["Availability"]);
                if (mo["BlockSize"] != null) fTreeView.Nodes[h].Nodes[u].Nodes[w].Nodes.Add("Размер блока: " + mo["BlockSize"]);
                if (mo["Bootable"] != null) fTreeView.Nodes[h].Nodes[u].Nodes[w].Nodes.Add("Загрузочный : " + mo["Bootable"]);

                if (mo["BootPartition"] != null) fTreeView.Nodes[h].Nodes[u].Nodes[w].Nodes.Add("Загрузочный раздел :" + mo["BootPartition"]);
                if (mo["Caption"] != null) fTreeView.Nodes[h].Nodes[u].Nodes[w].Nodes.Add("Заголовок : " + mo["Caption"]);
                if (mo["ConfigManagerErrorCode"] != null) fTreeView.Nodes[h].Nodes[u].Nodes[w].Nodes.Add("Код ошибки конфигурации: " + mo["ConfigManagerErrorCode"]);
                if (mo["ConfigManagerUserConfig"] != null) fTreeView.Nodes[h].Nodes[u].Nodes[w].Nodes.Add("Конфигурация пользователя: " + mo["ConfigManagerUserConfig"]);
                if (mo["CreationClassName"] != null) fTreeView.Nodes[h].Nodes[u].Nodes[w].Nodes.Add("Имя класса при создании: " + mo["CreationClassName"]);
                if (mo["Description"] != null) fTreeView.Nodes[h].Nodes[u].Nodes[w].Nodes.Add("Описание : " + mo["Description"]);
                if (mo["DeviceID"] != null) fTreeView.Nodes[h].Nodes[u].Nodes[w].Nodes.Add("Идентификатор устройства : " + mo["DeviceID"]);
                if (mo["DiskIndex"] != null) fTreeView.Nodes[h].Nodes[u].Nodes[w].Nodes.Add("Индекс диска : " + mo["DiskIndex"]);
                if (mo["ErrorCleared"] != null) fTreeView.Nodes[h].Nodes[u].Nodes[w].Nodes.Add("Ошибка при очистке диска : " + mo["ErrorCleared"]);
                if (mo["ErrorDescription"] != null) fTreeView.Nodes[h].Nodes[u].Nodes[w].Nodes.Add("Дескриптор ошибки: " + mo["ErrorDescription"]);
                if (mo["ErrorMethodology"] != null) fTreeView.Nodes[h].Nodes[u].Nodes[w].Nodes.Add("Тип ошибки: " + mo["ErrorMethodology"]);
                if (mo["HiddenSectors"] != null) fTreeView.Nodes[h].Nodes[u].Nodes[w].Nodes.Add("Количество скрытых секторов :  " + mo["HiddenSectors"]);
                if (mo["Index"] != null) fTreeView.Nodes[h].Nodes[u].Nodes[w].Nodes.Add("Индекс : " + mo["Index"]);
                if (mo["InstallDate"] != null) fTreeView.Nodes[h].Nodes[u].Nodes[w].Nodes.Add("Дата инсталляции : " + mo["InstallDate"]);
                if (mo["LastErrorCode"] != null) fTreeView.Nodes[h].Nodes[u].Nodes[w].Nodes.Add("Код последней ошибки : " + mo["LastErrorCode"]);
                if (mo["NumberOfBlocks"] != null) fTreeView.Nodes[h].Nodes[u].Nodes[w].Nodes.Add("Количество блоков : " + mo["NumberOfBlocks"]);
                if (mo["PNPDeviceID"] != null) fTreeView.Nodes[h].Nodes[u].Nodes[w].Nodes.Add("Сведения о Plug and Play: " + mo["PNPDeviceID"]);
                if (mo["PowerManagementCapabilities"] != null)
                {
                    v = 0;
                    UInt16[] arrPowerManagementCapabilities = (UInt16[])(mo["PowerManagementCapabilities"]);
                    foreach (UInt16 arrValue in arrPowerManagementCapabilities)
                    {
                        fTreeView.Nodes[h].Nodes[u].Nodes[w].Nodes[v].Nodes.Add("PowerManagementCapabilities: " + arrValue);
                        v++;
                    }
                }
                if (mo["PowerManagementSupported"] != null) fTreeView.Nodes[h].Nodes[u].Nodes[w].Nodes.Add("Поддержка управления питанием : " + mo["PowerManagementSupported"]);
                if (mo["PrimaryPartition"] != null) fTreeView.Nodes[h].Nodes[u].Nodes[w].Nodes.Add("Первичный раздел : " + mo["PrimaryPartition"]);
                w++;
            }
            u++;
        }

        //------------------------------------------------------
        // Получение учетных записей локальной машины
        //------------------------------------------------------
        public void UserAccount(TreeView aTreeView)
        {
            fTreeView = aTreeView;
            fTreeView.Nodes[h].Nodes.Add("Учетные записи");
            fTreeView.Nodes[h].Nodes[u].ImageIndex = 19;
            WqlObjectQuery query = new WqlObjectQuery("SELECT * FROM Win32_UserAccount WHERE LocalAccount='true'");
            ManagementObjectSearcher find = new ManagementObjectSearcher(query);
            int w = 0;
            foreach (ManagementObject mo in find.Get())
            {
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Имя: " + mo["Name"]);
                if (mo["AccountType"] != null)
                {
                    fTreeView.Nodes[h].Nodes[u].Nodes[w].ImageIndex = 15;
                    fTreeView.Nodes[h].Nodes[u].Nodes[w].Nodes.Add("AccountType: " + mo["AccountType"]);
                }
                if (mo["Caption"] != null)
                {
                    fTreeView.Nodes[h].Nodes[u].Nodes[w].ImageIndex = 15;
                    fTreeView.Nodes[h].Nodes[u].Nodes[w].Nodes.Add("Подпись : " + mo["Caption"]);
                }
                if (mo["Description"] != null) fTreeView.Nodes[h].Nodes[u].Nodes[w].Nodes.Add("Описание: " + mo["Description"]);
                if (mo["Disabled"] != null)
                    if (mo["Disabled"].ToString() == "true") fTreeView.Nodes[h].Nodes[u].Nodes[w].Nodes.Add("Доступна :  Да");
                    else fTreeView.Nodes[h].Nodes[u].Nodes[w].Nodes.Add("Доступна :  Нет");
                if (mo["Domain"] != null) fTreeView.Nodes[h].Nodes[u].Nodes[w].Nodes.Add("Домен : " + mo["Domain"]);
                if (mo["InstallDate"] != null) fTreeView.Nodes[h].Nodes[u].Nodes[w].Nodes.Add("InstallDate: " + mo["InstallDate"]);
                if (mo["LocalAccount"] != null) fTreeView.Nodes[h].Nodes[u].Nodes[w].Nodes.Add("LocalAccount: " + mo["LocalAccount"]);
                if (mo["Lockout"] != null) fTreeView.Nodes[h].Nodes[u].Nodes[w].Nodes.Add("Lockout: " + mo["Lockout"]);
                if (mo["FullName"] != null) fTreeView.Nodes[h].Nodes[u].Nodes[w].Nodes.Add("FullName: " + mo["FullName"]);
                if (mo["PasswordChangeable"] != null) fTreeView.Nodes[h].Nodes[u].Nodes[w].Nodes.Add("PasswordChangeable: " + mo["PasswordChangeable"]);
                if (mo["PasswordExpires"] != null) fTreeView.Nodes[h].Nodes[u].Nodes[w].Nodes.Add("PasswordExpires: " + mo["PasswordExpires"]);
                if (mo["PasswordRequired"] != null) fTreeView.Nodes[h].Nodes[u].Nodes[w].Nodes.Add("PasswordRequired: " + mo["PasswordRequired"]);
                if (mo["SID"] != null) fTreeView.Nodes[h].Nodes[u].Nodes[w].Nodes.Add("SID: " + mo["SID"]);
                // if (mo["SIDType"] != null) fTreeView.Nodes[h].Nodes[u].Nodes[w].Nodes.Add("SIDType: " + GetSidType(Convert.ToInt32(mo["SIDType"])));
                w++;
            }
            u++;
        }
        //------------------------------------------------------
        // Получение списка групп локальной машины
        //------------------------------------------------------
        public void Group(TreeView aTreeView)
        {
            fTreeView = aTreeView;
            fTreeView.Nodes[h].Nodes.Add("Список групп локальной машины");
            fTreeView.Nodes[h].Nodes[u].ImageIndex = 20;

            WqlObjectQuery query = new WqlObjectQuery("SELECT * FROM Win32_Group WHERE LocalAccount='true'");
            ManagementObjectSearcher find = new ManagementObjectSearcher(query);
            int w = 0;
            foreach (ManagementObject mo in find.Get())
            {
                fTreeView.Nodes[h].Nodes[u].Nodes.Add("Имя:   " + mo["Name"]);
                if (mo["Caption"] != null) fTreeView.Nodes[h].Nodes[u].Nodes[w].Nodes.Add("Название:   " + mo["Caption"]);
                if (mo["Description"] != null) fTreeView.Nodes[h].Nodes[u].Nodes[w].Nodes.Add("Описание:   " + mo["Description"]);
                if (mo["Domain"] != null) fTreeView.Nodes[h].Nodes[u].Nodes[w].Nodes.Add("Домен:    " + mo["Domain"]);
                if (mo["InstallDate"] != null) fTreeView.Nodes[h].Nodes[u].Nodes[w].Nodes.Add("Дата инсталляции: " + mo["InstallDate"]);
                if (mo["LocalAccount"] != null) fTreeView.Nodes[h].Nodes[u].Nodes[w].Nodes.Add("Локальная учетная запись: " + mo["LocalAccount"]);
                if (mo["SID"] != null) fTreeView.Nodes[h].Nodes[u].Nodes[w].Nodes.Add("SID:    " + mo["SID"]);
                // if (mo["SIDType"] != null) fTreeView.Nodes[h].Nodes[u].Nodes[w].Nodes.Add("Тип SID:    " + GetSidType(Convert.ToInt32(mo["SIDType"])));
                if (mo["Status"] != null) fTreeView.Nodes[h].Nodes[u].Nodes[w].Nodes.Add("Status: " + mo["Status"]);
                w++;
            }
            u++;
        }

        //---------------------------------------------
        // Получение информации о системе в целом
        //---------------------------------------------
        public void WMIAllInfo(int ah, System.Management.ManagementScope asc)
        {
            //fTreeView = aTreeView;
            sc = asc;
            h = ah;
            fTreeView.ImageIndex = 15;
            fTreeView.Nodes[h].ImageIndex = 0;

            foreach (String str in fCheckedListBox.CheckedItems)
            {
                if (str == "Свойства видеоконтроллера") VideoColntroller(fTreeView);//7
                if (str == "Информация о типе компьютера") ComputerSystem(fTreeView); //8
                if (str == "Информация о фирме-изготовителе") ComputerSystemProduct(fTreeView);//10
                if (str == "Информация об операционной системе") OperatingSystem(fTreeView);  //9
                if (str == "Информация о типе компьютера") ComputerSystem_1(fTreeView); //11
                if (str == "Физические параметры компьютера") SystemEnclosure(fTreeView); //12
                if (str == "Информация о процессорах") Processor(fTreeView); //14
                if (str == "Список общих файлов и каталогов") ShareFilesAndDirect(fTreeView); //5
                if (str == "Список подключенных сетевых ресурсов") NetworkConnection(fTreeView); //4
                //if (str == "Информация о диске") LogicalDisk(fTreeView);  //1
                if (str == "Информация о загрузчике") BootCofiguration(fTreeView);//3
                if (str == "Список запущенных сервисов") Process1(fTreeView);//13
                if (str == "Список процессов") Process(fTreeView);  //17
                if (str == "Информация о разделах диска") DiskPartition(fTreeView);  //18
                if (str == "Список учетных записей пользователей") UserAccount(fTreeView);  //19
                if (str == "Список групп") Group(fTreeView);  //20

            }
        }
    }
    //---------------------------------------------
    // Класс удаленные компьютеры
    //---------------------------------------------
    public class WMIInfoRemoteHost : WMIInfoLocalHost
    {
        public DataGridView dataGridView;
        public TextBox textBox;
        public CheckedListBox logCheckedListBox;
        public ComboBox ComboBox;
        //---------------------------------------------
        // Получение информации о системе в целом
        //---------------------------------------------
        public void WMIAllInfoRemoteHost(TreeView aTreeView, ComboBox acomboBox, CheckedListBox alogcheckedListBox, CheckedListBox acheckedListBox, DataGridView adataGridView, TextBox atextBox, int ah)
        {
            fTreeView = aTreeView;
            logCheckedListBox = alogcheckedListBox;
            fCheckedListBox = acheckedListBox;
            dataGridView = adataGridView;
            textBox = atextBox;
            ComboBox = acomboBox;
            string machine_name = "";
            string login = "";
            string pass = "";
            string ErrorMessage = "";
            string tmp = "";
            int i = 0;
            Boolean flg = false;
            textBox.Clear();
            int h = 0;
            fTreeView.Nodes.Clear();
            Boolean flag = false;
            MessageBox.Show("Вывод обновленной информации о системе", "Перезагрузка", MessageBoxButtons.OK, MessageBoxIcon.Information);
            System.Management.ManagementScope scope;
            // Просмотр списка выделенных компьютеров
            foreach (string str in logCheckedListBox.CheckedItems)
            {
                // Проверка возможности доступа к информации с помощью технологии WMI
                if ((textBox.Text == "Сервер Apple") || (textBox.Text == "Сервер Novell") || (textBox.Text == "Компьтеры с ОС LanManaget, входящие в состав домена") ||
                    (textBox.Text == "Cетевые принтеры") || (textBox.Text == "Xenix сервер") || (textBox.Text == "Сервер Unix"))
                {
                    textBox.Text += str + ". Технология WMI не может использоваться для получения сведений о данном компьютере " + (char)13 + (char)10;
                }
                else
                {
                    // Поиск логина и пароля в таблице
                    i = 0;
                    flg = false;
                    while (!flg && i < dataGridView.RowCount - 1)
                    {
                        tmp = dataGridView[0, i].Value.ToString();
                        if (str.ToUpper() == tmp.ToUpper()) flg = true;
                        else i++;
                    }
                    // Если компьютер найден,извлекается логин и пароль
                    if (flg)
                    {
                        flag = true;
                        ConnectionOptions options = new ConnectionOptions();
                        machine_name = dataGridView[0, i].Value.ToString();
                        login = dataGridView[1, i].Value.ToString();
                        login = login.Trim();
                        pass = dataGridView[2, i].Value.ToString();
                        pass = pass.Trim();

                        //if ((login.Length==0)&&(pass.Length==0))
                        if (1 == 1)
                            scope = new System.Management.ManagementScope("\\\\" + machine_name.ToUpper() + "\\root\\cimv2");
                        else
                        {
                            // Проверка соединения
                            options.Username = login;
                            options.Password = pass;
                            scope = new ManagementScope("\\\\" + machine_name.ToUpper() + "\\root\\cimv2", options);
                        }

                    }
                    else
                    {
                        // Сообщение об ошибке
                        ErrorMessage = "Компьютер " + str + " проверьте корректность ввода имени компьютера";
                        textBox.Text += ErrorMessage + (char)13 + (char)10;
                    }
                }
            }
            if (!flag)
            {
                ErrorMessage = "Ни один компьютер не выбран. Выводится информация о локальном компьютере";
                textBox.Clear();
                textBox.Text += ErrorMessage + (char)13 + (char)10;
            }

        }
    }
    //------------------------------------------------------------------------------------
    // Просмотр журналов аудита
    //------------------------------------------------------------------------------------
    public class AuditLog
    {
        public TreeView logTreeView;
        public TreeNode InfoNode;
        public DataGridView logDataGridView, errorDataGridView;
        public TextBox logTextBox;
        public int logFilterNumber;
        public DateTimePicker logDateTimePicker1;
        public DateTimePicker logDateTimePicker2;
        public GroupBox groupBox;
        //---------------------------------------------
        // Получение списка доступных журналов
        //---------------------------------------------
        public void EventLogList(TreeView aTreeView)
        {
            // Создание дерева журналов
            logTreeView = aTreeView;
            InfoNode = new TreeNode();
            InfoNode.Text = "Список доступных журналов аудита";
            logTreeView.Nodes.Add(InfoNode);
            EventLog[] elogs = EventLog.GetEventLogs();
            foreach (EventLog elog in elogs)
            {
                logTreeView.Nodes[0].Nodes.Add(elog.Log.ToString() + " (" + elog.LogDisplayName.ToString() + ")");
                elog.Close();
            }
            elogs = null;
        }
        //---------------------------------------------
        // Просмотр выбранного журнала
        //---------------------------------------------
        public void EventLogSee(TreeView aTreeView, DataGridView aDataGridView, GroupBox agroupBox)
        {

            logDataGridView = aDataGridView;
            errorDataGridView = aDataGridView;
            logDataGridView.Rows.Clear();
            errorDataGridView.Rows.Clear();
            groupBox = agroupBox;
            string tmp = "";
            string logType = "";
            int k = 0;
            if (logTreeView.SelectedNode != null)
            {
                tmp = logTreeView.SelectedNode.Text;
                k = tmp.IndexOf(" (");
                if (k != 0) logType = tmp.Substring(0, k);
                EventLog ev = new EventLog(logType, System.Environment.MachineName);
                int LastLogToShow = ev.Entries.Count;
                if (LastLogToShow <= 0)
                {
                    string message = "В журнале: " + logType + " отсутствуют записи";
                    string caption = "Просмотр журналов событий";
                    MessageBox.Show(message, caption, MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    // чтение выбранного журнала
                    groupBox.Text = "Просмотр журнала " + logType;
                    int i = 0;
                    for (i = 0; i < LastLogToShow; i++)
                    {
                        EventLogEntry CurrentEntry = ev.Entries[i];
                        logDataGridView.Rows.Add(CurrentEntry.EventID, CurrentEntry.EntryType.ToString(), CurrentEntry.TimeGenerated.ToString(), CurrentEntry.MachineName.ToString(), CurrentEntry.Message);

                    }
                }

                ev.Close();
            }
            else MessageBox.Show("Ни один журнал не выбран", "Просмотр журналов событий", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        //---------------------------------------------
        // Просмотр выбранного журнала с учетом фильтра
        //---------------------------------------------
        public void EventLogSeeFilter(TreeView aTreeView, DataGridView aDataGridView, TextBox aTextBox, int aFilterNumber)
        {
            // Передача переменных
            logTreeView = aTreeView;
            logDataGridView = aDataGridView;
            logTextBox = aTextBox;
            logFilterNumber = aFilterNumber;

            logDataGridView.Rows.Clear();
            string tmp = "";
            string logType = "";
            int k = 0;
            if (logTreeView.SelectedNode != null)
            {
                tmp = logTreeView.SelectedNode.Text;
                k = tmp.IndexOf(" (");
                if (k != 0) logType = tmp.Substring(0, k);
                EventLog ev = new EventLog(logType, System.Environment.MachineName);
                int LastLogToShow = ev.Entries.Count;
                if (LastLogToShow <= 0)
                {
                    string message = "В журнале: " + logType + " отсутствуют записи";
                    string caption = "Просмотр журналов событий";
                    MessageBox.Show(message, caption, MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    // чтение выбранного журнала
                    int i = 0;

                    for (i = 0; i < LastLogToShow; i++)
                    {
                        EventLogEntry CurrentEntry = ev.Entries[i];
                        switch (logFilterNumber)
                        {
                            case 0:
                            {
                                if (string.Equals(logTextBox.Text, CurrentEntry.EventID.ToString()))
                                    logDataGridView.Rows.Add(CurrentEntry.EventID, CurrentEntry.EntryType.ToString(), CurrentEntry.TimeGenerated.ToString(), CurrentEntry.MachineName.ToString(), CurrentEntry.Message);
                                break;
                            }
                            case 1:
                                if (string.Equals(logTextBox.Text, CurrentEntry.EntryType.ToString()))
                                    logDataGridView.Rows.Add(CurrentEntry.EventID, CurrentEntry.EntryType.ToString(), CurrentEntry.TimeGenerated.ToString(), CurrentEntry.MachineName.ToString(), CurrentEntry.Message);
                                break;

                            case 3:
                                if (string.Equals(logTextBox.Text, CurrentEntry.MachineName.ToString()))
                                    logDataGridView.Rows.Add(CurrentEntry.EventID, CurrentEntry.EntryType.ToString(), CurrentEntry.TimeGenerated.ToString(), CurrentEntry.MachineName.ToString(), CurrentEntry.Message);
                                break;
                        }
                    }
                }
                ev.Close();
            }
            else MessageBox.Show("Ни один журнал не выбран", "Просмотр журналов событий", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        //---------------------------------------------
        // Просмотр выбранного журнала с учетом фильтра
        //---------------------------------------------
        public void EventLogSeeFilter(TreeView aTreeView, DataGridView aDataGridView, DateTimePicker aDateTimePicker1, DateTimePicker aDateTimePicker2)
        {
            // Передача переменных
            logTreeView = aTreeView;
            logDataGridView = aDataGridView;
            logDateTimePicker1 = aDateTimePicker1;
            logDateTimePicker2 = aDateTimePicker2;
            logDataGridView.Rows.Clear();
            DateTimePicker tmpdateTimePicker;
            string tmp = "";
            string logType = "";
            int k = 0;
            if (logTreeView.SelectedNode != null)
            {
                tmp = logTreeView.SelectedNode.Text;
                k = tmp.IndexOf(" (");
                if (k != 0) logType = tmp.Substring(0, k);
                EventLog ev = new EventLog(logType, System.Environment.MachineName);
                int LastLogToShow = ev.Entries.Count;
                if (LastLogToShow <= 0)
                {
                    string message = "В журнале: " + logType + " отсутствуют записи";
                    string caption = "Просмотр журналов событий";
                    MessageBox.Show(message, caption, MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    // чтение выбранного журнала
                    int i = 0;

                    for (i = 0; i < LastLogToShow; i++)
                    {
                        EventLogEntry CurrentEntry = ev.Entries[i];
                        if (logDateTimePicker1.Value > logDateTimePicker2.Value)
                        {
                            tmpdateTimePicker = logDateTimePicker1;
                            logDateTimePicker1 = logDateTimePicker2;
                            logDateTimePicker2 = tmpdateTimePicker;
                        }
                        if ((CurrentEntry.TimeGenerated.Date >= logDateTimePicker1.Value) && (CurrentEntry.TimeGenerated.Date <= logDateTimePicker2.Value))
                            logDataGridView.Rows.Add(CurrentEntry.EventID, CurrentEntry.EntryType.ToString(), CurrentEntry.TimeGenerated.ToString(), CurrentEntry.MachineName.ToString(), CurrentEntry.Message);
                    }
                }
                ev.Close();
            }
            else MessageBox.Show("Ни один журнал не выбран", "Просмотр журналов событий", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        //---------------------------------------------
        // Очистка выбранного журнала
        //---------------------------------------------
        public void EventLogClear(TreeView aTreeView, DataGridView aDataGridView)
        {
            logDataGridView = aDataGridView;
            logDataGridView.Rows.Clear();
            string tmp = "";
            string logType = "";
            int k = 0;
            if (logTreeView.SelectedNode != null)
            {
                tmp = logTreeView.SelectedNode.Text;
                k = tmp.IndexOf(" (");
                if (k != 0) logType = tmp.Substring(0, k);
                EventLog ev = new EventLog(logType, System.Environment.MachineName);
                int LastLogToShow = ev.Entries.Count;
                if (LastLogToShow <= 0)
                {
                    string message = "В журнале: " + logType + " записи уже отсутствуют";
                    string caption = "Очистка журналов событий";
                    MessageBox.Show(message, caption, MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    ev.Clear();
                    ev.Close();
                }
            }
            else MessageBox.Show("Ни один журнал не выбран", "Очистка журналов событий", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public void ErrorLogSee(TreeView aTreeView, DataGridView aDataGridView, GroupBox agroupBox)
        {

            logDataGridView = aDataGridView;
            errorDataGridView = aDataGridView;
            logDataGridView.Rows.Clear();
            errorDataGridView.Rows.Clear();
            groupBox = agroupBox;
            string tmp = "";
            string logType = "";
            int k = 0;
            if (logTreeView.SelectedNode != null)
            {
                tmp = logTreeView.SelectedNode.Text;
                k = tmp.IndexOf(" (");
                if (k != 0) logType = tmp.Substring(0, k);
                EventLog ev = new EventLog(logType, System.Environment.MachineName);
                int LastLogToShow = ev.Entries.Count;
                if (LastLogToShow <= 0)
                {
                    string message = "В журнале: " + logType + " отсутствуют записи";
                    string caption = "Просмотр журналов событий";
                    MessageBox.Show(message, caption, MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    // чтение выбранного журнала
                    groupBox.Text = "Просмотр журнала " + logType;
                    int i = 0;
                    int itog = 0;
                    int warning = 0;
                    int error = 0;


                    for (i = 0; i < LastLogToShow; i++)
                    {
                        EventLogEntry CurrentEntry = ev.Entries[i];
                        //switch  (CurrentEntry.EventID)
                        // {
                        //     case : logDataGridView.Rows.Add(CurrentEntry.EventID, "Ошибка", CurrentEntry.TimeGenerated.ToString(), CurrentEntry.MachineName.ToString(), CurrentEntry.Message);
                        //         itog++;
                        //          break;         
                        //  }
                        if ((CurrentEntry.EventID == 4) || (CurrentEntry.EventID == 472) || (CurrentEntry.EventID == 477) || (CurrentEntry.EventID == 517) ||
                            (CurrentEntry.EventID == 624) || (CurrentEntry.EventID == 535) || (CurrentEntry.EventID == 533) || (CurrentEntry.EventID == 529) ||
                            (CurrentEntry.EventID == 632) || (CurrentEntry.EventID == 539) || (CurrentEntry.EventID == 534) || (CurrentEntry.EventID == 531) ||
                            (CurrentEntry.EventID == 636) || (CurrentEntry.EventID == 660) || (CurrentEntry.EventID == 6806) || (CurrentEntry.EventID == 4645) ||
                            (CurrentEntry.EventID == 642) || (CurrentEntry.EventID == 675) || (CurrentEntry.EventID == 681) || (CurrentEntry.EventID == 4728) ||
                            (CurrentEntry.EventID == 644) || (CurrentEntry.EventID == 676) || (CurrentEntry.EventID == 1102) || (CurrentEntry.EventID == 4732) ||
                            (CurrentEntry.EventID == 4740) || (CurrentEntry.EventID == 4756) || (CurrentEntry.EventID == 4768) || (CurrentEntry.EventID == 4776) ||
                            (CurrentEntry.EventID == 4738) || (CurrentEntry.EventID == 4733) || (CurrentEntry.EventID == 630))

                        {
                            logDataGridView.Rows.Add(CurrentEntry.EventID, "Обратите внимание", CurrentEntry.TimeGenerated.ToString(), CurrentEntry.MachineName.ToString(), CurrentEntry.Message);
                            itog++;
                            warning++;

                        }
                        if ((CurrentEntry.EventID == 200) || (CurrentEntry.EventID == 4124) || (CurrentEntry.EventID == 4226) || (CurrentEntry.EventID == 7901) ||
                            (CurrentEntry.EventID == 12294) || (CurrentEntry.EventID == 9095) || (CurrentEntry.EventID == 9097) || (CurrentEntry.EventID == 7023) ||
                            (CurrentEntry.EventID == 6183) || (CurrentEntry.EventID == 55) || (CurrentEntry.EventID == 1066) || (CurrentEntry.EventID == 6008) || (CurrentEntry.EventID == 861))
                        {
                            logDataGridView.Rows.Add(CurrentEntry.EventID, "ОШИБКА!!!", CurrentEntry.TimeGenerated.ToString(), CurrentEntry.MachineName.ToString(), CurrentEntry.Message);
                            itog++;
                            error++;
                        }



                    }
                    if (itog == 0)
                        MessageBox.Show("Опасных записей в журнале не найдено", "Система отслеживания вторжений", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    else MessageBox.Show("Всего найдено   " + itog + "   записей" + "\r\n" + "Из них   " + warning + "        с пометкой 'Обратите внимание'" + "\r\n" + "              " + error + "        с пометкой 'ОШИБКА!!!'",
                        "Результат проверки", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                ev.Close();
            }
            else MessageBox.Show("Ни один журнал не выбран", "Просмотр журналов событий", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }


    }

}


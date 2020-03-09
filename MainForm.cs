using System;
using System.Windows.Forms;
using System.Collections;
using System.Web.Script.Serialization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

namespace EventHook.WinForms.Example
{

    public partial class MainForm : Form
    {
        
        private readonly ApplicationWatcher applicationWatcher;
        private readonly ClipboardWatcher clipboardWatcher;
        private readonly EventHookFactory eventHookFactory = new EventHookFactory();
        private readonly KeyboardWatcher keyboardWatcher;
        private readonly MouseWatcher mouseWatcher;
        private readonly PrintWatcher printWatcher;

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")]
        private static extern IntPtr GetActiveWindow();
        [DllImport("user32.dll")]
        private static extern IntPtr GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);
        [DllImport("user32.dll", EntryPoint = "WindowFromPoint")]//position to get form       
        public static extern IntPtr WindowFromPoint(int xPoint, int yPoint);
        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        public ArrayList eventlog = new ArrayList();
        public string job_id = Guid.NewGuid().ToString();
        public string resource = "Human";
        public DateTime current_time = new DateTime();
        public string whole_log = "";
        public string[] keyboard_special_key = { "Tab", "LeftAlt", "LeftShift", "Capital", "LeftCtrl", "LWin", "Space", "RightAlt", "RWin", "Apps", "Escape", "RightCtrl", "RightShift", "Return", "Back", "Left", "Up", "Down", "Right", "Delete", "Insert", "Home", "PageUp", "Next", "End", "Snapshot", "Scroll", "Pause" };

        public MainForm()
        {

            Application.ApplicationExit += OnApplicationExit;
            InitializeComponent();

            keyboardWatcher = eventHookFactory.GetKeyboardWatcher();
            keyboardWatcher.OnKeyInput += (s, e) =>
            {
                
                if (e.KeyData.EventType.ToString() == "down") {
                    current_time = DateTime.Now;
                    Hashtable this_event = new Hashtable
                    {
                        { "Event Type", "Keyboard" },
                        { "Event Info", e.KeyData.Keyname },
                        { "Job id", job_id },
                        { "Resource", resource },
                        { "Form", GainActivatedForm() },
                        { "Timestamp", current_time.ToString("yyyy-MM-dd HH:mm:ss") +"."+ current_time.Millisecond.ToString()}
                    };

                    eventlog.Add(this_event);
                }
            };

            mouseWatcher = eventHookFactory.GetMouseWatcher();
            string[] MouseMoveMent = { "WM_WHEELBUTTONDOWN", "WM_LBUTTONDOWN", "WM_MOUSEWHEEL", "WM_WHEELBUTTONDOWN" };
            mouseWatcher.OnMouseInput += (s, e) =>
            {
                if (Array.IndexOf(MouseMoveMent, e.Message.ToString()) != -1)
                {
                    current_time = DateTime.Now;
                    string[] inter = GainMouseClickClassName(e.Point.x, e.Point.y);
                    Hashtable this_event = new Hashtable
                    {
                        { "Event Type", "Mouse" },
                        { "Event Info",  e.Message.ToString() },
                        { "Position", e.Point.x.ToString() + "|" + e.Point.y.ToString() },
                        { "Job id", job_id },
                        { "Resource", resource },
                        { "Form", GainActivatedForm()},
                        { "Click_Name", inter[1]},
                        { "Click_ClassName", inter[0]},
                        { "Timestamp", current_time.ToString("yyyy-MM-dd HH:mm:ss") +"."+ current_time.Millisecond.ToString()}
                    };

                    eventlog.Add(this_event);
                };
            };

            applicationWatcher = eventHookFactory.GetApplicationWatcher();
            applicationWatcher.OnApplicationWindowChange += (s, e) =>
            {
                Hashtable this_event = new Hashtable
                    {
                        { "Event Type", "Application" },
                        { "Event Info", e.ApplicationData.AppName + "|" + e.ApplicationData.AppTitle+ "|" +e.Event },
                        { "Job id", job_id },
                        { "Resource", resource },
                        { "Form", GainActivatedForm() },
                        { "Timestamp", current_time.ToString("yyyy-MM-dd HH:mm:ss") +"."+ current_time.Millisecond.ToString()}
                    };

                eventlog.Add(this_event);

                //Console.WriteLine("Application window of '{0}' with the title '{1}' was {2}",
                //    e.ApplicationData.AppName, e.ApplicationData.AppTitle, e.Event);
            };

            clipboardWatcher = eventHookFactory.GetClipboardWatcher();
            clipboardWatcher.OnClipboardModified += (s, e) =>
            {
                Hashtable this_event = new Hashtable
                    {
                        { "Event Type", "Clipboard" },
                        { "Event Info", e.Data },
                        { "Job id", job_id },
                        { "Resource", resource },
                        { "Form", GainActivatedForm() },
                        { "Timestamp", current_time.ToString("yyyy-MM-dd HH:mm:ss") +"."+ current_time.Millisecond.ToString()}
                    };

                eventlog.Add(this_event);

                //Console.WriteLine("Clipboard updated with data '{0}' of format {1}", e.Data,
                //    e.DataFormat.ToString());
            };


            //////////////////////////////////////////////////////////////////
            printWatcher = eventHookFactory.GetPrintWatcher();
            printWatcher.OnPrintEvent += (s, e) =>
            {
                Hashtable this_event = new Hashtable
                    {
                        { "Event Type", "Printer" },
                        { "Event Info", e.EventData.PrinterName},
                        { "Job id", job_id },
                        { "Resource", resource },
                        { "Form", GainActivatedForm() },
                        { "Timestamp", current_time.ToString("yyyy-MM-dd HH:mm:ss") +"."+ current_time.Millisecond.ToString()}
                    };

                eventlog.Add(this_event);

                //Console.WriteLine("Printer '{0}' currently printing {1} pages.", e.EventData.PrinterName,
                //    e.EventData.Pages);
            };
        }

        private String GainActivatedForm()
        {
            IntPtr myPtr = GetForegroundWindow();
            StringBuilder str = new StringBuilder(512);
            GetWindowText(GetForegroundWindow(), str, str.Capacity);
            //GetWindowText(GetActiveWindow(), str, str.Capacity);

            return str.ToString();
        }

        private string[] GainMouseClickClassName(int x, int y)
        {
            IntPtr hWnd = WindowFromPoint(x, y);
            StringBuilder WindowText = new StringBuilder(512);
            GetWindowText(hWnd, WindowText, WindowText.Capacity);

            StringBuilder ClassName = new StringBuilder(512);
            GetClassName(hWnd, ClassName, ClassName.Capacity);

            string[] inter = {ClassName.ToString(), WindowText.ToString() };
            return inter;
        }

        private void OnApplicationExit(object sender, EventArgs e)
        {
            keyboardWatcher.Stop();
            mouseWatcher.Stop();
            clipboardWatcher.Stop();
            applicationWatcher.Stop();
            printWatcher.Stop();

            eventHookFactory.Dispose();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            job_id = Guid.NewGuid().ToString();
            label3.Text = job_id;
        }

        private void button2_Click(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e) //start
        {
            //MessageBox.Show(resource);
            keyboardWatcher.Start();
            mouseWatcher.Start();
            clipboardWatcher.Start();
            applicationWatcher.Start();
            printWatcher.Start();

            label4.Text = "Recording";

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void RadioButton1_CheckedChanged(object sender, EventArgs e)
        {
            resource = "Human";
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            resource = "Robot";
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click_2(object sender, EventArgs e) // put into event log
        {
            string filePath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + "\\Action_Event_Log.json"; // desktop location
            FileStream fs = new FileStream(filePath, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            fs.SetLength(0);
            StreamWriter sw = new StreamWriter(fs);
            whole_log = "[" + whole_log.Substring(1) + "]";
            sw.Write(whole_log);
            sw.Close();

            label4.Text = "Finish";
            whole_log = "";

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e) // stop
        {
            if (eventlog.Count > 0)
            {

                try
                {
                    keyboardWatcher.Stop();
                    mouseWatcher.Stop();
                    clipboardWatcher.Stop();
                    applicationWatcher.Stop();
                    printWatcher.Stop();
                }
                catch {
                    MessageBox.Show("Watcher fail");
                    eventHookFactory.Dispose();
                    button3.Enabled = false;
                }

                ArrayList inter_eventlog = new ArrayList();
                Hashtable inter_keyboard_event = new Hashtable();
                string inter_keyboard_contain = "";

                foreach (Hashtable i in eventlog)
                {

                    if (i["Event Type"].ToString() == "Keyboard" && Array.IndexOf(keyboard_special_key, i["Event Info"]) == -1)
                    {
                        inter_keyboard_contain += i["Event Info"];
                        inter_keyboard_event = i;
                    }
                    else
                    {
                        if (inter_keyboard_contain != "")
                        {
                            inter_keyboard_event["Event Info"] = inter_keyboard_contain;
                            inter_keyboard_contain = "";
                            inter_eventlog.Add(inter_keyboard_event);
                        }
                        inter_eventlog.Add(i);
                    }

                }

                var serializer = new JavaScriptSerializer();
                var serializedResult = serializer.Serialize(inter_eventlog);
                serializedResult = serializedResult.Replace("[", "");
                serializedResult = serializedResult.Replace("]", "");
                whole_log += "," + serializedResult;
                eventlog.Clear();
                MessageBox.Show(inter_eventlog.Count.ToString() + " records");

                label4.Text = "Holding Log";

            }
            else
            {
                    label4.Text = "No Log";
            }

        }


    }
}


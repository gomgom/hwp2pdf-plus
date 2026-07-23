// hwp2pdf+ (HWP to PDF Plus) — HWP/HWPX -> PDF 일괄 변환기
// 한/글 오토메이션을 늦은 바인딩(IDispatch)으로 호출하므로 한/글 버전이 바뀌어도 깨지지 않는다.
// (기존 hwp2pdf.exe는 구버전 타입 라이브러리 조기 바인딩이라 한/글 2024에서 첫 호출부터 죽는다.)
//
// 빌드: csc.exe /target:winexe /codepage:65001 /win32icon:res\app.ico /resource:res\app.ico,app.ico /out:hwp2pdf-plus.exe Program.cs
// CLI 모드: hwp2pdf-plus.exe <파일|폴더> ...  (GUI 없이 변환, 결과는 원본 옆에 생성)
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace Hwp2PdfPlus
{
    // 늦은 바인딩 래퍼: HWPFrame.HwpObject 를 IDispatch 로만 호출
    class HwpAutomation : IDisposable
    {
        private object _hwp;
        private readonly Type _type;

        public HwpAutomation()
        {
            Type t = Type.GetTypeFromProgID("HWPFrame.HwpObject");
            if (t == null)
                throw new InvalidOperationException("한/글(한컴오피스)이 설치되어 있지 않습니다.");
            _type = t;
            _hwp = Activator.CreateInstance(t);
            Invoke("RegisterModule", "FilePathCheckDLL", "FilePathCheckerModuleExample");
        }

        private object Invoke(string name, params object[] args)
        {
            return _type.InvokeMember(name, BindingFlags.InvokeMethod, null, _hwp, args);
        }

        public bool Convert(string src, string pdf)
        {
            bool opened = (bool)Invoke("Open", src, "", "forceopen:true");
            if (!opened) return false;
            try
            {
                return (bool)Invoke("SaveAs", pdf, "PDF", "");
            }
            finally
            {
                Invoke("Run", "FileClose");
            }
        }

        public void Dispose()
        {
            if (_hwp != null)
            {
                try { Invoke("Quit"); } catch { }
                try { Marshal.ReleaseComObject(_hwp); } catch { }
                _hwp = null;
            }
        }
    }

    class MainForm : Form
    {
        private readonly ListView _list = new ListView();
        private readonly TextBox _log = new TextBox();
        private readonly Button _btnAdd = new Button();
        private readonly Button _btnClear = new Button();
        private readonly Button _btnConvert = new Button();
        private readonly CheckBox _chkOverwrite = new CheckBox();
        private bool _busy;

        public MainForm()
        {
            Text = "hwp2pdf+ — HWP → PDF 일괄 변환기";
            AllowDrop = true;
            MinimumSize = new Size(560, 420);
            Size = new Size(680, 520);

            try
            {
                using (var s = GetType().Assembly.GetManifestResourceStream("app.ico"))
                    if (s != null) Icon = new Icon(s);
            }
            catch { }

            _list.View = View.Details;
            _list.FullRowSelect = true;
            _list.Columns.Add("파일명", 200);
            _list.Columns.Add("경로", 260);
            _list.Columns.Add("상태", 120);
            _list.Dock = DockStyle.Fill;
            _list.AllowDrop = true;
            _list.DragEnter += OnDragEnter;
            _list.DragDrop += OnDragDrop;
            _list.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Delete && !_busy)
                    foreach (ListViewItem it in _list.SelectedItems.Cast<ListViewItem>().ToList())
                        it.Remove();
            };

            _log.Multiline = true;
            _log.ReadOnly = true;
            _log.ScrollBars = ScrollBars.Vertical;
            _log.Dock = DockStyle.Bottom;
            _log.Height = 110;

            var panel = new FlowLayoutPanel();
            panel.Dock = DockStyle.Top;
            panel.Height = 36;
            panel.Padding = new Padding(4);

            _btnAdd.Text = "파일 추가";
            _btnAdd.Click += (s, e) => AddViaDialog();
            _btnClear.Text = "목록 비우기";
            _btnClear.Click += (s, e) => { if (!_busy) _list.Items.Clear(); };
            _btnConvert.Text = "PDF 변환";
            _btnConvert.Click += async (s, e) => await ConvertAllAsync();
            _chkOverwrite.Text = "기존 PDF 덮어쓰기";
            _chkOverwrite.AutoSize = true;
            _chkOverwrite.Margin = new Padding(12, 8, 0, 0);

            foreach (var b in new[] { _btnAdd, _btnClear, _btnConvert })
            {
                b.AutoSize = true;
                b.AutoSizeMode = AutoSizeMode.GrowAndShrink;
                b.Padding = new Padding(8, 2, 8, 2);
                b.Margin = new Padding(0, 4, 6, 4);
            }

            panel.Controls.Add(_btnAdd);
            panel.Controls.Add(_btnClear);
            panel.Controls.Add(_btnConvert);
            panel.Controls.Add(_chkOverwrite);

            Controls.Add(_list);
            Controls.Add(panel);
            Controls.Add(_log);

            DragEnter += OnDragEnter;
            DragDrop += OnDragDrop;

            Log("HWP/HWPX 파일이나 폴더를 창에 끌어다 놓은 뒤 [PDF 변환]을 누르세요. PDF는 원본 옆에 생성됩니다.");
        }

        private void OnDragEnter(object sender, DragEventArgs e)
        {
            if (!_busy && e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.Copy;
        }

        private void OnDragDrop(object sender, DragEventArgs e)
        {
            if (_busy) return;
            AddPaths((string[])e.Data.GetData(DataFormats.FileDrop));
        }

        private void AddViaDialog()
        {
            using (var dlg = new OpenFileDialog())
            {
                dlg.Filter = "한/글 문서 (*.hwp;*.hwpx)|*.hwp;*.hwpx|모든 파일 (*.*)|*.*";
                dlg.Multiselect = true;
                if (dlg.ShowDialog(this) == DialogResult.OK)
                    AddPaths(dlg.FileNames);
            }
        }

        private void AddPaths(IEnumerable<string> paths)
        {
            var files = new List<string>();
            foreach (string p in paths)
            {
                if (Directory.Exists(p))
                    files.AddRange(Directory.EnumerateFiles(p, "*.*", SearchOption.AllDirectories)
                        .Where(IsHwpFile));
                else if (File.Exists(p) && IsHwpFile(p))
                    files.Add(p);
            }
            var existing = new HashSet<string>(
                _list.Items.Cast<ListViewItem>().Select(i => (string)i.Tag),
                StringComparer.OrdinalIgnoreCase);
            int added = 0;
            foreach (string f in files.Where(f => !existing.Contains(f)))
            {
                var item = new ListViewItem(new[] { Path.GetFileName(f), Path.GetDirectoryName(f), "대기" });
                item.Tag = f;
                _list.Items.Add(item);
                existing.Add(f);
                added++;
            }
            if (added > 0) Log(added + "개 파일을 추가했습니다. (총 " + _list.Items.Count + "개)");
        }

        private static bool IsHwpFile(string path)
        {
            string ext = Path.GetExtension(path);
            return ext.Equals(".hwp", StringComparison.OrdinalIgnoreCase)
                || ext.Equals(".hwpx", StringComparison.OrdinalIgnoreCase);
        }

        private async System.Threading.Tasks.Task ConvertAllAsync()
        {
            if (_busy || _list.Items.Count == 0) return;
            _busy = true;
            _btnConvert.Enabled = _btnAdd.Enabled = _btnClear.Enabled = false;
            bool overwrite = _chkOverwrite.Checked;
            var items = _list.Items.Cast<ListViewItem>().ToList();
            int ok = 0, fail = 0, skip = 0;

            try
            {
                await System.Threading.Tasks.Task.Run(() =>
                {
                    using (var hwp = new HwpAutomation())
                    {
                        foreach (var item in items)
                        {
                            string src = (string)item.Tag;
                            string pdf = Path.ChangeExtension(src, ".pdf");
                            SetStatus(item, "변환 중...");
                            try
                            {
                                if (!overwrite && File.Exists(pdf))
                                {
                                    SetStatus(item, "건너뜀(PDF 있음)");
                                    skip++;
                                    continue;
                                }
                                if (hwp.Convert(src, pdf))
                                {
                                    SetStatus(item, "완료");
                                    ok++;
                                }
                                else
                                {
                                    SetStatus(item, "실패");
                                    fail++;
                                }
                            }
                            catch (Exception ex)
                            {
                                SetStatus(item, "오류");
                                fail++;
                                LogAsync("오류: " + Path.GetFileName(src) + " — " + ex.Message);
                            }
                        }
                    }
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "한/글 연결에 실패했습니다.\n" + ex.Message, "오류",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                _busy = false;
                _btnConvert.Enabled = _btnAdd.Enabled = _btnClear.Enabled = true;
                Log("변환 종료 — 완료 " + ok + ", 실패 " + fail + ", 건너뜀 " + skip);
            }
        }

        private void SetStatus(ListViewItem item, string status)
        {
            if (InvokeRequired) BeginInvoke((Action)(() => item.SubItems[2].Text = status));
            else item.SubItems[2].Text = status;
        }

        private void LogAsync(string msg)
        {
            if (InvokeRequired) BeginInvoke((Action)(() => Log(msg)));
            else Log(msg);
        }

        private void Log(string msg)
        {
            _log.AppendText("[" + DateTime.Now.ToString("HH:mm:ss") + "] " + msg + Environment.NewLine);
        }
    }

    static class Program
    {
        [STAThread]
        static int Main(string[] args)
        {
            if (args.Length > 0)
                return RunCli(args);

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
            return 0;
        }

        // CLI 모드: 인자로 받은 파일/폴더를 GUI 없이 변환
        static int RunCli(string[] args)
        {
            var files = new List<string>();
            foreach (string p in args)
            {
                if (Directory.Exists(p))
                    files.AddRange(Directory.EnumerateFiles(p, "*.*", SearchOption.AllDirectories)
                        .Where(f => f.EndsWith(".hwp", StringComparison.OrdinalIgnoreCase)
                                 || f.EndsWith(".hwpx", StringComparison.OrdinalIgnoreCase)));
                else if (File.Exists(p))
                    files.Add(Path.GetFullPath(p));
                else
                    Console.Error.WriteLine("찾을 수 없음: " + p);
            }
            if (files.Count == 0)
            {
                Console.Error.WriteLine("변환할 .hwp/.hwpx 파일이 없습니다.");
                return 1;
            }
            int fail = 0;
            using (var hwp = new HwpAutomation())
            {
                foreach (string src in files)
                {
                    string pdf = Path.ChangeExtension(src, ".pdf");
                    try
                    {
                        if (hwp.Convert(src, pdf))
                            Console.WriteLine("OK   " + src);
                        else { Console.WriteLine("FAIL " + src); fail++; }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("FAIL " + src + " — " + ex.Message);
                        fail++;
                    }
                }
            }
            return fail == 0 ? 0 : 2;
        }
    }
}

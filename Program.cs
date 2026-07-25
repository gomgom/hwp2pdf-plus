// hwp2pdf+ (HWP to PDF Plus) — HWP/HWPX -> PDF 일괄 변환기
// 한/글 오토메이션을 늦은 바인딩(IDispatch)으로 호출하므로 한/글 버전이 바뀌어도 깨지지 않는다.
// (기존 hwp2pdf.exe는 구버전 타입 라이브러리 조기 바인딩이라 한/글 2024에서 첫 호출부터 죽는다.)
//
// 화면은 WPF. ui\MainWindow.xaml 을 리소스로 임베드해 XamlReader로 런타임 파싱하므로
// 코드비하인드 컴파일 없이 csc.exe 단일 빌드가 유지된다. 빌드는 build.ps1 참조.
// CLI 모드: hwp2pdf-plus.exe <파일|폴더> ...  (GUI 없이 변환, 결과는 원본 옆에 생성)
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;

[assembly: AssemblyTitle("hwp2pdf+")]
[assembly: AssemblyProduct("hwp2pdf+ (HWP to PDF Plus)")]
[assembly: AssemblyDescription("HWP/HWPX → PDF 일괄 변환기")]
[assembly: AssemblyCompany("gomgom")]
[assembly: AssemblyCopyright("Copyright (c) 2026 gomgom")]
[assembly: AssemblyVersion("1.1.0.0")]
[assembly: AssemblyFileVersion("1.1.0.0")]

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

        // 임의의 COM 객체(액션·파라미터 셋 등)에 대한 늦은 바인딩 호출
        private static object InvokeOn(object obj, string name, params object[] args)
        {
            return obj.GetType().InvokeMember(name, BindingFlags.InvokeMethod, null, obj, args);
        }

        // useSavedSettings=true 이면 한/글에 저장된 마지막 PDF 설정(변환 범위 등)을 그대로 사용한다.
        // false(기본)이면 잔존 설정을 무시하고 항상 문서 전체를 출력한다.
        public bool Convert(string src, string pdf, bool useSavedSettings)
        {
            bool opened = (bool)Invoke("Open", src, "", "forceopen:true");
            if (!opened) return false;
            try
            {
                // PrintToPDFEx가 기존 파일에 덮어쓰기 프롬프트를 띄우지 않도록 미리 삭제
                try { if (File.Exists(pdf)) File.Delete(pdf); } catch { }

                // 저장된 설정 사용: SaveAs 필터가 한/글에 저장된 마지막 변환 범위·모아찍기를 물려받는다.
                if (useSavedSettings)
                    return (bool)Invoke("SaveAs", pdf, "PDF", "");

                // 기본: 가상 인쇄(PrintToPDFEx) + 기본값 로드(GetDefault)로 항상 문서 전체를 출력한다.
                // (SaveAs 필터는 잔존 설정 때문에 일부 페이지만 나올 수 있어 기본 경로로 쓰지 않는다.)
                if (SaveAsPdfFullDocument(pdf)) return true;

                // 폴백: PrintToPDFEx 실패 시(예: "Hancom PDF" 프린터 부재) 기존 방식으로라도 저장
                return (bool)Invoke("SaveAs", pdf, "PDF", "");
            }
            finally
            {
                Invoke("Run", "FileClose");
            }
        }

        private bool SaveAsPdfFullDocument(string pdf)
        {
            try
            {
                object act = Invoke("CreateAction", "PrintToPDFEx");
                object set = InvokeOn(act, "CreateSet");
                InvokeOn(act, "GetDefault", set);
                InvokeOn(set, "SetItem", "PrinterName", "Hancom PDF");
                InvokeOn(set, "SetItem", "FileName", pdf);
                InvokeOn(set, "SetItem", "PrintMethod", 0); // 모아찍기(N-up) 해제
                object r = InvokeOn(act, "Execute", set);
                return r is bool && (bool)r && File.Exists(pdf);
            }
            catch
            {
                return false;
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

    // 목록 한 줄 (XAML DataTemplate 이 바인딩한다)
    public class FileItem : INotifyPropertyChanged
    {
        private string _status = "대기";
        private string _kind = "idle";

        public FileItem(string path)
        {
            FullPath = path;
            Name = Path.GetFileName(path);
            Folder = Path.GetDirectoryName(path);
        }

        public string FullPath { get; private set; }
        public string Name { get; private set; }
        public string Folder { get; private set; }
        public string Status { get { return _status; } }
        public string StatusKind { get { return _kind; } }   // idle | busy | done | fail | skip

        public void Set(string status, string kind)
        {
            _status = status;
            _kind = kind;
            Raise("Status");
            Raise("StatusKind");
        }

        public event PropertyChangedEventHandler PropertyChanged;
        private void Raise(string n)
        {
            var h = PropertyChanged;
            if (h != null) h(this, new PropertyChangedEventArgs(n));
        }
    }

    class MainUi
    {
        public readonly Window Win;

        private readonly ListBox _list;
        private readonly Button _btnConvert, _btnAdd, _btnClear, _btnPick;
        private readonly CheckBox _chkOverwrite, _chkUseSaved, _chkOutFolder;
        private readonly TextBlock _txtOutPath, _txtStatus, _txtCount;
        private readonly ProgressBar _bar;
        private readonly TextBox _log;
        private readonly UIElement _spinner, _empty, _dropOverlay;
        private readonly ObservableCollection<FileItem> _items = new ObservableCollection<FileItem>();
        private readonly DispatcherTimer _dragTimer = new DispatcherTimer();
        private readonly Microsoft.Win32.UserPreferenceChangedEventHandler _themeHandler;

        private string _outDir;
        private bool _busy;

        public MainUi()
        {
            Win = LoadXaml();

            _list = (ListBox)Win.FindName("FileList");
            _btnConvert = (Button)Win.FindName("BtnConvert");
            _btnAdd = (Button)Win.FindName("BtnAdd");
            _btnClear = (Button)Win.FindName("BtnClear");
            _btnPick = (Button)Win.FindName("BtnPickFolder");
            _chkOverwrite = (CheckBox)Win.FindName("ChkOverwrite");
            _chkUseSaved = (CheckBox)Win.FindName("ChkUseSaved");
            _chkOutFolder = (CheckBox)Win.FindName("ChkOutFolder");
            _txtOutPath = (TextBlock)Win.FindName("TxtOutPath");
            _txtStatus = (TextBlock)Win.FindName("TxtStatus");
            _txtCount = (TextBlock)Win.FindName("TxtCount");
            _bar = (ProgressBar)Win.FindName("Bar");
            _log = (TextBox)Win.FindName("Log");
            _spinner = (UIElement)Win.FindName("Spinner");
            _empty = (UIElement)Win.FindName("EmptyState");
            _dropOverlay = (UIElement)Win.FindName("DropOverlay");

            _list.ItemsSource = _items;
            _items.CollectionChanged += (s, e) => UpdateEmptyState();

            SetWindowIcon();
            ApplyTheme();
            Win.SourceInitialized += (s, e) => ApplyTitleBarTheme();

            _themeHandler = (s, e) =>
            {
                if (e.Category == Microsoft.Win32.UserPreferenceCategory.General ||
                    e.Category == Microsoft.Win32.UserPreferenceCategory.Color ||
                    e.Category == Microsoft.Win32.UserPreferenceCategory.VisualStyle)
                    Win.Dispatcher.BeginInvoke((Action)(() => { ApplyTheme(); ApplyTitleBarTheme(); }));
            };
            Microsoft.Win32.SystemEvents.UserPreferenceChanged += _themeHandler;
            Win.Closed += (s, e) => Microsoft.Win32.SystemEvents.UserPreferenceChanged -= _themeHandler;

            // ── 이벤트 연결 ──
            _btnConvert.Click += async (s, e) => await ConvertAllAsync();
            _btnAdd.Click += (s, e) => AddViaDialog();
            _btnClear.Click += (s, e) => { if (!_busy) _items.Clear(); };
            _btnPick.Click += (s, e) => PickOutputFolder();
            _chkOutFolder.Checked += (s, e) => UpdateOutputFolderState();
            _chkOutFolder.Unchecked += (s, e) => UpdateOutputFolderState();

            _list.KeyDown += (s, e) =>
            {
                if (e.Key == Key.Delete && !_busy)
                    foreach (var it in _list.SelectedItems.Cast<FileItem>().ToList())
                        _items.Remove(it);
            };

            // 드래그&드롭 — 오버레이는 DragOver 가 끊기면(=창 밖으로 나가면) 타이머로 감춘다
            _dragTimer.Interval = TimeSpan.FromMilliseconds(220);
            _dragTimer.Tick += (s, e) => { _dragTimer.Stop(); ShowDropOverlay(false); };

            Win.PreviewDragEnter += OnDragOver;
            Win.PreviewDragOver += OnDragOver;
            Win.PreviewDrop += OnDrop;

            UpdateEmptyState();
            UpdateOutputFolderState();
            Log("HWP/HWPX 파일이나 폴더를 창에 끌어다 놓은 뒤 [PDF 변환]을 누르세요. 출력 폴더를 선택하지 않는 경우, 원본 파일과 같은 폴더에 생성됩니다.");
        }

        // ─────────────────────────── 화면 로드 · 테마

        private static Window LoadXaml()
        {
            using (var s = typeof(MainUi).Assembly.GetManifestResourceStream("MainWindow.xaml"))
            {
                if (s == null) throw new InvalidOperationException("MainWindow.xaml 리소스를 찾을 수 없습니다.");
                return (Window)XamlReader.Load(s);
            }
        }

        private void SetWindowIcon()
        {
            try
            {
                using (var s = typeof(MainUi).Assembly.GetManifestResourceStream("app.ico"))
                {
                    if (s == null) return;
                    try
                    {
                        var dec = new IconBitmapDecoder(s, BitmapCreateOptions.None, BitmapCacheOption.OnLoad);
                        BitmapFrame best = null;
                        foreach (var f in dec.Frames)
                            if (best == null || f.PixelWidth > best.PixelWidth) best = f;
                        Win.Icon = best;
                    }
                    catch
                    {
                        s.Position = 0;
                        Win.Icon = BitmapFrame.Create(s, BitmapCreateOptions.None, BitmapCacheOption.OnLoad);
                    }
                }
            }
            catch { }
        }

        private static bool IsDarkTheme()
        {
            try
            {
                using (var k = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(
                    @"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"))
                {
                    if (k == null) return false;
                    object v = k.GetValue("AppsUseLightTheme");
                    return v is int && (int)v == 0;
                }
            }
            catch { return false; }
        }

        private static Brush B(string hex)
        {
            var b = new SolidColorBrush((Color)ColorConverter.ConvertFromString(hex));
            b.Freeze();
            return b;
        }

        private void ApplyTheme()
        {
            bool dark = IsDarkTheme();
            var r = Win.Resources;

            // 강조색 — 앱 아이콘의 HWP 문서 파랑 계열
            r["Accent"] = B(dark ? "#3B8DDD" : "#1E6FBF");
            r["AccentHover"] = B(dark ? "#529CE6" : "#2A7ED0");
            r["AccentPressed"] = B(dark ? "#2F7CC7" : "#17589A");

            r["WindowBg"] = B(dark ? "#16181C" : "#F3F4F6");
            r["CardBg"] = B(dark ? "#1E2126" : "#FFFFFF");
            r["BorderSoft"] = B(dark ? "#2C3038" : "#E4E6EA");
            r["Fg"] = B(dark ? "#E9EBEE" : "#1A1C1E");
            r["FgMuted"] = B(dark ? "#A8AEB8" : "#5A6069");
            r["FgSubtle"] = B(dark ? "#7C838E" : "#8A9099");
            r["InputBg"] = B(dark ? "#191C21" : "#FAFBFC");

            r["GhostBg"] = B(dark ? "#252931" : "#FFFFFF");
            r["GhostHover"] = B(dark ? "#2E333C" : "#F2F4F7");
            r["GhostPressed"] = B(dark ? "#363C46" : "#E8EBEF");
            r["DisabledBg"] = B(dark ? "#2A2E35" : "#DDE1E6");

            r["TrackBg"] = B(dark ? "#2E333B" : "#E6E8EC");
            r["TrackBorder"] = B(dark ? "#3A404A" : "#D0D4DA");
            r["ThumbFill"] = B(dark ? "#9AA1AC" : "#8A9099");
            r["ScrollThumb"] = B(dark ? "#3E444D" : "#C7CBD1");

            r["RowHover"] = B(dark ? "#262A31" : "#F2F4F7");
            r["RowSelected"] = B(dark ? "#1E3A57" : "#E6F0FA");
            r["DropBorder"] = B(dark ? "#444B55" : "#C3C9D1");
            r["DropOverlayBg"] = B(dark ? "#E01A1D22" : "#E6F4F8FC");

            r["IdleBg"] = B(dark ? "#2A2F36" : "#EDEFF2");
            r["IdleFg"] = B(dark ? "#A8AEB8" : "#5A6069");
            r["BusyBg"] = B(dark ? "#16324D" : "#E3EFFB");
            r["BusyFg"] = B(dark ? "#79B8F3" : "#1A5FA6");
            r["OkBg"] = B(dark ? "#12331F" : "#E4F6EA");
            r["OkFg"] = B(dark ? "#74D69B" : "#1B7A3E");
            r["ErrBg"] = B(dark ? "#3A1B1B" : "#FDE8E8");
            r["ErrFg"] = B(dark ? "#F09A9A" : "#B02525");
            r["WarnBg"] = B(dark ? "#3A2E15" : "#FFF4E0");
            r["WarnFg"] = B(dark ? "#E3B667" : "#96631A");
        }

        [DllImport("dwmapi.dll", PreserveSig = true)]
        private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int value, int size);

        private void ApplyTitleBarTheme()
        {
            try
            {
                IntPtr h = new WindowInteropHelper(Win).Handle;
                if (h == IntPtr.Zero) return;
                int on = IsDarkTheme() ? 1 : 0;
                if (DwmSetWindowAttribute(h, 20, ref on, sizeof(int)) != 0) // DWMWA_USE_IMMERSIVE_DARK_MODE
                    DwmSetWindowAttribute(h, 19, ref on, sizeof(int));      // 구버전 빌드용
            }
            catch { }
        }

        // ─────────────────────────── 목록 · 옵션

        private void UpdateEmptyState()
        {
            _empty.Visibility = _items.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
            if (!_busy) _txtCount.Text = _items.Count > 0 ? _items.Count + "개" : "";
        }

        private void UpdateOutputFolderState()
        {
            bool on = _chkOutFolder.IsChecked == true;
            _btnPick.IsEnabled = on && !_busy;
            if (!on)
            {
                _outDir = null;
                _txtOutPath.Text = "선택된 폴더 없음";
            }
            else if (string.IsNullOrEmpty(_outDir))
            {
                _txtOutPath.Text = "폴더를 선택하세요";
            }
        }

        private void PickOutputFolder()
        {
            string picked = FolderPicker.Pick(Win, _outDir);
            if (string.IsNullOrEmpty(picked)) return;
            _outDir = picked;
            _txtOutPath.Text = picked;
            Log("출력 폴더: " + picked);
        }

        private void AddViaDialog()
        {
            var dlg = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "한/글 문서 (*.hwp;*.hwpx)|*.hwp;*.hwpx|모든 파일 (*.*)|*.*",
                Multiselect = true
            };
            if (dlg.ShowDialog(Win) == true) AddPaths(dlg.FileNames);
        }

        private void AddPaths(IEnumerable<string> paths)
        {
            var files = new List<string>();
            foreach (string p in paths)
            {
                if (Directory.Exists(p))
                    files.AddRange(Directory.EnumerateFiles(p, "*.*", SearchOption.AllDirectories).Where(IsHwpFile));
                else if (File.Exists(p) && IsHwpFile(p))
                    files.Add(p);
            }
            var existing = new HashSet<string>(_items.Select(i => i.FullPath), StringComparer.OrdinalIgnoreCase);
            int added = 0;
            foreach (string f in files.Where(f => !existing.Contains(f)))
            {
                _items.Add(new FileItem(f));
                existing.Add(f);
                added++;
            }
            if (added > 0) Log(added + "개 파일을 추가했습니다. (총 " + _items.Count + "개)");
        }

        private static bool IsHwpFile(string path)
        {
            string ext = Path.GetExtension(path);
            return ext.Equals(".hwp", StringComparison.OrdinalIgnoreCase)
                || ext.Equals(".hwpx", StringComparison.OrdinalIgnoreCase);
        }

        // ─────────────────────────── 드래그 & 드롭

        private static bool HasFiles(DragEventArgs e)
        {
            return e.Data != null && e.Data.GetDataPresent(DataFormats.FileDrop);
        }

        private void OnDragOver(object sender, DragEventArgs e)
        {
            bool ok = !_busy && HasFiles(e);
            e.Effects = ok ? DragDropEffects.Copy : DragDropEffects.None;
            e.Handled = true;
            if (ok)
            {
                ShowDropOverlay(true);
                _dragTimer.Stop();
                _dragTimer.Start();
            }
        }

        private void OnDrop(object sender, DragEventArgs e)
        {
            _dragTimer.Stop();
            ShowDropOverlay(false);
            e.Handled = true;
            if (_busy || !HasFiles(e)) return;
            AddPaths((string[])e.Data.GetData(DataFormats.FileDrop));
        }

        private void ShowDropOverlay(bool on)
        {
            _dropOverlay.Visibility = on ? Visibility.Visible : Visibility.Collapsed;
        }

        // ─────────────────────────── 변환

        private void SetBusy(bool busy, int total)
        {
            _busy = busy;
            _btnConvert.IsEnabled = _btnAdd.IsEnabled = _btnClear.IsEnabled = !busy;
            _chkOverwrite.IsEnabled = _chkUseSaved.IsEnabled = _chkOutFolder.IsEnabled = !busy;
            _btnPick.IsEnabled = !busy && _chkOutFolder.IsChecked == true;
            _spinner.Visibility = busy ? Visibility.Visible : Visibility.Collapsed;

            // 변환 중에는 한/글 창에 가리지 않도록 항상 위로
            Win.Topmost = busy;
            if (busy)
            {
                Win.Activate();
            }
            else
            {
                // 변환이 끝나면 포커스를 이 창으로 되돌린다.
                // 한/글(Hwp.exe)이 백그라운드에서 뜨면서 포그라운드를 가져가기 때문에,
                // Topmost 를 해제하는 순간 다른 창이 앞으로 나온 것처럼 보인다.
                Win.Dispatcher.BeginInvoke((Action)(() =>
                {
                    if (Win.WindowState == WindowState.Minimized) return;
                    Win.Activate();
                    Win.Focus();
                }), DispatcherPriority.ApplicationIdle);
            }

            _bar.Maximum = total > 0 ? total : 1;
            _bar.Value = 0;
            if (!busy)
            {
                _txtStatus.Text = "준비됨";
                UpdateEmptyState();
            }
        }

        private async System.Threading.Tasks.Task ConvertAllAsync()
        {
            if (_busy || _items.Count == 0) return;

            string outDir = null;
            if (_chkOutFolder.IsChecked == true)
            {
                if (string.IsNullOrEmpty(_outDir))
                {
                    MessageBox.Show(Win, "출력 폴더를 선택하세요.", "출력 폴더",
                        MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }
                outDir = _outDir;
            }
            bool overwrite = _chkOverwrite.IsChecked == true;
            bool useSaved = _chkUseSaved.IsChecked == true;

            var items = _items.ToList();
            SetBusy(true, items.Count);
            foreach (var it in items) it.Set("대기", "idle");
            _txtStatus.Text = "변환 준비 중...";

            int ok = 0, fail = 0, skip = 0;
            try
            {
                await System.Threading.Tasks.Task.Run(() =>
                {
                    using (var hwp = new HwpAutomation())
                    {
                        for (int i = 0; i < items.Count; i++)
                        {
                            var item = items[i];
                            int idx = i + 1;
                            string src = item.FullPath;
                            string pdf = outDir != null
                                ? Path.Combine(outDir, Path.GetFileNameWithoutExtension(src) + ".pdf")
                                : Path.ChangeExtension(src, ".pdf");

                            UI(() =>
                            {
                                item.Set("변환 중", "busy");
                                _txtStatus.Text = "변환 중 (" + idx + "/" + items.Count + "): " + item.Name;
                                _txtCount.Text = idx + " / " + items.Count;
                                _list.ScrollIntoView(item);
                            });

                            try
                            {
                                if (!overwrite && File.Exists(pdf))
                                {
                                    skip++;
                                    UI(() => item.Set("건너뜀", "skip"));
                                }
                                else if (hwp.Convert(src, pdf, useSaved))
                                {
                                    ok++;
                                    UI(() => item.Set("완료", "done"));
                                }
                                else
                                {
                                    fail++;
                                    UI(() => item.Set("실패", "fail"));
                                }
                            }
                            catch (Exception ex)
                            {
                                fail++;
                                string msg = ex.Message;
                                UI(() =>
                                {
                                    item.Set("오류", "fail");
                                    Log("오류: " + item.Name + " — " + msg);
                                });
                            }

                            UI(() => _bar.Value = idx);
                        }
                    }
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show(Win, "한/글 연결에 실패했습니다.\n\n" + ex.Message, "오류",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                SetBusy(false, items.Count);
                Log("변환 종료 — 완료 " + ok + ", 실패 " + fail + ", 건너뜀 " + skip);
            }
        }

        private void UI(Action a)
        {
            if (Win.Dispatcher.CheckAccess()) a();
            else Win.Dispatcher.Invoke(a);
        }

        private void Log(string msg)
        {
            _log.AppendText("[" + DateTime.Now.ToString("HH:mm:ss") + "] " + msg + Environment.NewLine);
            _log.ScrollToEnd();
        }
    }

    // 모던 폴더 선택창(IFileOpenDialog). 실패하면 구형 대화상자로 폴백한다.
    static class FolderPicker
    {
        private const uint FOS_PICKFOLDERS = 0x00000020;
        private const uint FOS_FORCEFILESYSTEM = 0x00000040;
        private const uint SIGDN_FILESYSPATH = 0x80058000;

        public static string Pick(Window owner, string initial)
        {
            try { return PickModern(owner, initial); }
            catch { try { return PickLegacy(initial); } catch { return null; } }
        }

        private static string PickModern(Window owner, string initial)
        {
            var dlg = (IFileOpenDialog)new FileOpenDialogRCW();
            uint opts;
            dlg.GetOptions(out opts);
            dlg.SetOptions(opts | FOS_PICKFOLDERS | FOS_FORCEFILESYSTEM);
            dlg.SetTitle("PDF를 저장할 폴더를 선택하세요");

            if (!string.IsNullOrEmpty(initial) && Directory.Exists(initial))
            {
                try
                {
                    Guid iid = typeof(IShellItem).GUID;
                    object si;
                    if (SHCreateItemFromParsingName(initial, IntPtr.Zero, ref iid, out si) == 0)
                        dlg.SetFolder((IShellItem)si);
                }
                catch { }
            }

            IntPtr hwnd = owner == null ? IntPtr.Zero : new WindowInteropHelper(owner).Handle;
            if (dlg.Show(hwnd) != 0) return null;   // 취소 포함

            IShellItem result;
            dlg.GetResult(out result);
            IntPtr p;
            result.GetDisplayName(SIGDN_FILESYSPATH, out p);
            string path = Marshal.PtrToStringUni(p);
            Marshal.FreeCoTaskMem(p);
            return path;
        }

        private static string PickLegacy(string initial)
        {
            using (var d = new System.Windows.Forms.FolderBrowserDialog())
            {
                d.Description = "PDF를 저장할 폴더를 선택하세요.";
                if (!string.IsNullOrEmpty(initial) && Directory.Exists(initial)) d.SelectedPath = initial;
                return d.ShowDialog() == System.Windows.Forms.DialogResult.OK ? d.SelectedPath : null;
            }
        }

        [DllImport("shell32.dll", CharSet = CharSet.Unicode, PreserveSig = true)]
        private static extern int SHCreateItemFromParsingName(
            [MarshalAs(UnmanagedType.LPWStr)] string pszPath, IntPtr pbc, ref Guid riid,
            [MarshalAs(UnmanagedType.Interface)] out object ppv);

        [ComImport, Guid("DC1C5A9C-E88A-4dde-A5A1-60F82A20AEF7")]
        private class FileOpenDialogRCW { }

        // 주의: COM vtable 순서대로 선언해야 한다(사용하지 않는 메서드도 자리를 지켜야 함).
        [ComImport, Guid("d57c7288-d4ad-4768-be02-9d969532d960"),
         InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IFileOpenDialog
        {
            // IModalWindow
            [PreserveSig] int Show(IntPtr parent);
            // IFileDialog
            void SetFileTypes(uint cFileTypes, IntPtr rgFilterSpec);
            void SetFileTypeIndex(uint iFileType);
            void GetFileTypeIndex(out uint piFileType);
            void Advise(IntPtr pfde, out uint pdwCookie);
            void Unadvise(uint dwCookie);
            void SetOptions(uint fos);
            void GetOptions(out uint pfos);
            void SetDefaultFolder(IShellItem psi);
            void SetFolder(IShellItem psi);
            void GetFolder(out IShellItem ppsi);
            void GetCurrentSelection(out IShellItem ppsi);
            void SetFileName([MarshalAs(UnmanagedType.LPWStr)] string pszName);
            void GetFileName([MarshalAs(UnmanagedType.LPWStr)] out string pszName);
            void SetTitle([MarshalAs(UnmanagedType.LPWStr)] string pszTitle);
            void SetOkButtonLabel([MarshalAs(UnmanagedType.LPWStr)] string pszText);
            void SetFileNameLabel([MarshalAs(UnmanagedType.LPWStr)] string pszLabel);
            void GetResult(out IShellItem ppsi);
        }

        [ComImport, Guid("43826d1e-e718-42ee-bc55-a1e261c37bfe"),
         InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IShellItem
        {
            void BindToHandler(IntPtr pbc, ref Guid bhid, ref Guid riid, out IntPtr ppv);
            void GetParent(out IShellItem ppsi);
            void GetDisplayName(uint sigdnName, out IntPtr ppszName);
            void GetAttributes(uint sfgaoMask, out uint psfgaoAttribs);
            void Compare(IShellItem psi, uint hint, out int piOrder);
        }
    }

    static class Program
    {
        [STAThread]
        static int Main(string[] args)
        {
            if (args.Length > 0)
                return RunCli(args);

            try
            {
                var app = new Application();
                var ui = new MainUi();
                app.Run(ui.Win);
                return 0;
            }
            catch (Exception ex)
            {
                MessageBox.Show("화면을 초기화하지 못했습니다.\n\n" + ex, "hwp2pdf+ 오류",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                return 3;
            }
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
                        if (hwp.Convert(src, pdf, false))
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

using System;
using System.IO;
using System.IO.Compression;
using System.Net;
using System.Net.Sockets;
using System.Diagnostics;
using System.Threading;
using System.Reflection;
using System.Drawing;
using System.Windows.Forms;

namespace SysgenLoansApp
{
    class Program
    {
        private static HttpListener _listener;
        private static string _webRoot;
        private static int _port;
        private static Mutex _singleInstanceMutex;

        [STAThread]
        static void Main(string[] args)
        {
            bool isNewInstance;
            _singleInstanceMutex = new Mutex(true, "SysgenRPOLoansAppSingleInstanceMutex", out isNewInstance);

            _port = 48921;
            string appUrl = "http://127.0.0.1:" + _port + "/";

            if (!isNewInstance)
            {
                // App already running: just launch the window and bring to front
                LaunchAppWindow(appUrl);
                return;
            }

            try
            {
                string exeDir = AppDomain.CurrentDomain.BaseDirectory;
                string distDir = Path.Combine(exeDir, "dist");

                if (Directory.Exists(distDir) && File.Exists(Path.Combine(distDir, "index.html")))
                {
                    _webRoot = distDir;
                }
                else if (File.Exists(Path.Combine(exeDir, "index.html")))
                {
                    _webRoot = exeDir;
                }
                else
                {
                    // Extract embedded web assets for 100% standalone USB portable execution
                    string localAppWeb = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "SysgenLoansApp", "web_v1");
                    EnsureEmbeddedAssetsExtracted(localAppWeb);
                    _webRoot = localAppWeb;
                }

                // Check port availability
                _port = GetAvailablePort(48921);
                appUrl = "http://127.0.0.1:" + _port + "/";

                // Start embedded web server
                StartWebServer(_port, _webRoot);

                // Launch App Window
                LaunchAppWindow(appUrl);

                // System Tray Management
                NotifyIcon tray = new NotifyIcon();
                try
                {
                    string iconFile = Path.Combine(exeDir, "app_icon.ico");
                    if (File.Exists(iconFile))
                    {
                        tray.Icon = new Icon(iconFile);
                    }
                    else
                    {
                        tray.Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath);
                    }
                }
                catch
                {
                    tray.Icon = SystemIcons.Application;
                }

                tray.Text = "Sysgen RPO Loans Suite";
                tray.Visible = true;

                ContextMenu menu = new ContextMenu();
                menu.MenuItems.Add("Open Workspace", (s, e) => LaunchAppWindow(appUrl));
                menu.MenuItems.Add("-");
                menu.MenuItems.Add("Exit Application", (s, e) => {
                    tray.Visible = false;
                    tray.Dispose();
                    Application.Exit();
                });

                tray.ContextMenu = menu;
                tray.DoubleClick += (s, e) => LaunchAppWindow(appUrl);

                // Keep application alive
                Application.Run();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Error launching Sysgen RPO Loans App: " + ex.Message,
                    "Sysgen RPO Loans Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
            finally
            {
                if (_listener != null && _listener.IsListening)
                {
                    try { _listener.Stop(); } catch { }
                }
                if (_singleInstanceMutex != null)
                {
                    try { _singleInstanceMutex.ReleaseMutex(); } catch { }
                }
            }
        }

        private static void EnsureEmbeddedAssetsExtracted(string targetDir)
        {
            if (Directory.Exists(targetDir) && File.Exists(Path.Combine(targetDir, "index.html")))
            {
                return;
            }

            Directory.CreateDirectory(targetDir);
            var asm = Assembly.GetExecutingAssembly();
            using (var stream = asm.GetManifestResourceStream("app_assets.zip"))
            {
                if (stream != null)
                {
                    using (var archive = new ZipArchive(stream))
                    {
                        foreach (var entry in archive.Entries)
                        {
                            string destPath = Path.Combine(targetDir, entry.FullName);
                            if (string.IsNullOrEmpty(entry.Name))
                            {
                                Directory.CreateDirectory(destPath);
                            }
                            else
                            {
                                Directory.CreateDirectory(Path.GetDirectoryName(destPath));
                                entry.ExtractToFile(destPath, true);
                            }
                        }
                    }
                }
            }
        }

        private static int GetAvailablePort(int startingPort)
        {
            for (int p = startingPort; p < startingPort + 100; p++)
            {
                try
                {
                    TcpListener l = new TcpListener(IPAddress.Loopback, p);
                    l.Start();
                    l.Stop();
                    return p;
                }
                catch { }
            }
            return 5173;
        }

        private static void StartWebServer(int port, string rootDir)
        {
            _listener = new HttpListener();
            _listener.Prefixes.Add("http://127.0.0.1:" + port + "/");
            _listener.Start();

            ThreadPool.QueueUserWorkItem((o) =>
            {
                while (_listener != null && _listener.IsListening)
                {
                    try
                    {
                        var context = _listener.GetContext();
                        ThreadPool.QueueUserWorkItem((ctxObj) =>
                        {
                            HandleRequest((HttpListenerContext)ctxObj, rootDir);
                        }, context);
                    }
                    catch
                    {
                        break;
                    }
                }
            });
        }

        private static void HandleRequest(HttpListenerContext context, string rootDir)
        {
            try
            {
                string rawUrl = context.Request.Url.AbsolutePath;
                if (rawUrl == "/" || string.IsNullOrEmpty(rawUrl))
                {
                    rawUrl = "/index.html";
                }

                string relativePath = rawUrl.TrimStart('/').Replace('/', Path.DirectorySeparatorChar);
                string filePath = Path.Combine(rootDir, relativePath);

                if (!File.Exists(filePath))
                {
                    filePath = Path.Combine(rootDir, "index.html");
                }

                if (File.Exists(filePath))
                {
                    byte[] bytes = File.ReadAllBytes(filePath);
                    context.Response.ContentType = GetMimeType(filePath);
                    context.Response.ContentLength64 = bytes.Length;
                    context.Response.AddHeader("Access-Control-Allow-Origin", "*");
                    context.Response.OutputStream.Write(bytes, 0, bytes.Length);
                    context.Response.StatusCode = 200;
                }
                else
                {
                    context.Response.StatusCode = 404;
                }
            }
            catch
            {
                context.Response.StatusCode = 500;
            }
            finally
            {
                try { context.Response.OutputStream.Close(); } catch { }
            }
        }

        private static string GetMimeType(string path)
        {
            string ext = Path.GetExtension(path).ToLower();
            switch (ext)
            {
                case ".html": return "text/html; charset=utf-8";
                case ".js": return "application/javascript; charset=utf-8";
                case ".css": return "text/css; charset=utf-8";
                case ".json": return "application/json";
                case ".webmanifest": return "application/manifest+json";
                case ".png": return "image/png";
                case ".jpg":
                case ".jpeg": return "image/jpeg";
                case ".svg": return "image/svg+xml";
                case ".ico": return "image/x-icon";
                case ".xlsx": return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                case ".woff2": return "font/woff2";
                case ".woff": return "font/woff";
                case ".ttf": return "font/ttf";
                default: return "application/octet-stream";
            }
        }

        private static void LaunchAppWindow(string url)
        {
            string edgePath = FindBrowserExecutable("msedge.exe");
            if (!string.IsNullOrEmpty(edgePath))
            {
                StartBrowserProcess(edgePath, url);
                return;
            }

            string chromePath = FindBrowserExecutable("chrome.exe");
            if (!string.IsNullOrEmpty(chromePath))
            {
                StartBrowserProcess(chromePath, url);
                return;
            }

            Process.Start(new ProcessStartInfo(url) { UseShellExecute = true });
        }

        private static void StartBrowserProcess(string browserExe, string url)
        {
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            string profileDir = Path.Combine(appData, "SysgenLoansAppProfile");

            var psi = new ProcessStartInfo
            {
                FileName = browserExe,
                Arguments = "--app=\"" + url + "\" --window-size=1440,900 --user-data-dir=\"" + profileDir + "\"",
                UseShellExecute = false
            };

            Process.Start(psi);
        }

        private static string FindBrowserExecutable(string exeName)
        {
            string[] possiblePaths = new string[]
            {
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), @"Microsoft\Edge\Application\" + exeName),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), @"Microsoft\Edge\Application\" + exeName),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), @"Google\Chrome\Application\" + exeName),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), @"Google\Chrome\Application\" + exeName),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"Microsoft\Edge\Application\" + exeName),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"Google\Chrome\Application\" + exeName)
            };

            foreach (string p in possiblePaths)
            {
                if (File.Exists(p)) return p;
            }

            return null;
        }
    }
}

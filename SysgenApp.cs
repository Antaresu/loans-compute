using System;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Diagnostics;
using System.Threading;
using System.Text;

namespace SysgenLoansApp
{
    class Program
    {
        private static HttpListener _listener;
        private static string _webRoot;
        private static int _port;

        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                string exeDir = AppDomain.CurrentDomain.BaseDirectory;
                string distDir = Path.Combine(exeDir, "dist");
                _webRoot = Directory.Exists(distDir) ? distDir : exeDir;

                _port = GetAvailablePort(48921);

                // Start local embedded static server
                StartWebServer(_port, _webRoot);

                string appUrl = "http://127.0.0.1:" + _port + "/";

                // Launch Edge or Chrome in true Standalone App Window Mode
                Process appProc = LaunchAppWindow(appUrl);

                if (appProc != null)
                {
                    appProc.WaitForExit();
                }
                else
                {
                    // Fallback: open default browser
                    Process.Start(new ProcessStartInfo(appUrl) { UseShellExecute = true });
                    Thread.Sleep(5000);
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(
                    "Error launching Sysgen RPO Loans App: " + ex.Message,
                    "Sysgen RPO Loans Error",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error
                );
            }
            finally
            {
                if (_listener != null && _listener.IsListening)
                {
                    try { _listener.Stop(); } catch { }
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
                    // Fallback to index.html for SPA routing
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

        private static Process LaunchAppWindow(string url)
        {
            // Search for Microsoft Edge (Standard Chromium installed on Windows 10/11)
            string edgePath = FindBrowserExecutable("msedge.exe");
            if (!string.IsNullOrEmpty(edgePath))
            {
                return StartBrowserProcess(edgePath, url);
            }

            // Search for Chrome
            string chromePath = FindBrowserExecutable("chrome.exe");
            if (!string.IsNullOrEmpty(chromePath))
            {
                return StartBrowserProcess(chromePath, url);
            }

            return null;
        }

        private static Process StartBrowserProcess(string browserExe, string url)
        {
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            string profileDir = Path.Combine(appData, "SysgenLoansAppProfile");

            var psi = new ProcessStartInfo
            {
                FileName = browserExe,
                Arguments = "--app=\"" + url + "\" --window-size=1440,900 --user-data-dir=\"" + profileDir + "\"",
                UseShellExecute = false,
                CreateNoWindow = true
            };

            return Process.Start(psi);
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

using Microsoft.Win32;
using SkiaSharp.QrCode.Image;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using ThemeCommons.Controls;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using System.Linq;

namespace PowerPoint_Remote
{

    public partial class MainWindow : DefaultWindow
    {
        private static string GetIpWithGateway()
        {
            foreach (var adapter in NetworkInterface.GetAllNetworkInterfaces())
            {
                if (adapter.OperationalStatus != OperationalStatus.Up ||
                    adapter.NetworkInterfaceType == NetworkInterfaceType.Loopback ||
                    adapter.NetworkInterfaceType == NetworkInterfaceType.Tunnel)
                    continue;

                var props = adapter.GetIPProperties();
                var hasGateway = props?.GatewayAddresses?.Any(g =>
                    g.Address.AddressFamily == AddressFamily.InterNetwork) ?? false;

                if (hasGateway)
                {
                    var unicast = props.UnicastAddresses
                                      .FirstOrDefault(ip => ip.Address.AddressFamily == AddressFamily.InterNetwork);
                    if (unicast != null)
                        return unicast.Address.ToString();
                }
            }

            return "No valid gateway IP found.";
        }

        private Server server;
        public MainWindow()
        {
            InitializeComponent();
            server = new Server();
            IpDisplayText.Text = $"Device IP: http://{GetIpWithGateway()}:{Server.Port}";
            Closing += (s, e) => server.Dispose();

            // Load default s.ppts file if it exists
            string presetPath = @"C:\Users\first\Desktop\s.pptx";
            if (File.Exists(presetPath))
            {
                PPTPath.Text = presetPath;
                server.OpenPresentation(presetPath);
                BuildQrCode();
                Activate();
            }
        }

        private void BuildQrCode()
        {
            var stream = new MemoryStream();
            var qrcode = new QrCode("http://" + GetIpWithGateway() + ":" + Server.Port, new Vector2Slim(512, 512), SkiaSharp.SKEncodedImageFormat.Png);
            qrcode.GenerateImage(stream);
            QrImg.Source = BitmapFrame.Create(stream, BitmapCreateOptions.None, BitmapCacheOption.OnLoad);
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            var ofd = new OpenFileDialog
            {
                Filter = "PowerPoint Presentation (.ppt(x)) |*.ppt;*.pptx"
            };
            if (ofd.ShowDialog(this) != true) return;
            PPTPath.Text = ofd.FileName;
            server.OpenPresentation(PPTPath.Text);
            BuildQrCode();
            Activate();
        }

        private void Hyperlink_RequestNavigate(object sender, RequestNavigateEventArgs e) =>
            Process.Start(new ProcessStartInfo(e.Uri.AbsoluteUri) { UseShellExecute = true });
    }
}

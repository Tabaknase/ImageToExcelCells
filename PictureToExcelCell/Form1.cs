using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PictureToExcelCell
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        public void ClipBoardPicture(string filename, int size)
        {
            var decregex = @"\,\d{8}";


            var hexregex = @"#[\d|\w]{6}";
            var rgbregex = @"rgb\(\d{1,3},\s\d{1,3},\s\d{1,3}\)";
            var firstColor = Clipboard.GetData(DataFormats.Html);
            var stringOfColor = "";

            if (firstColor != null) stringOfColor = firstColor.ToString();
            string testColor =
                "Version:0.9\r\nStartHTML:0000000153\r\nEndHTML:0000000767\r\nStartFragment:0000000189\r\nEndFragment:0000000731\r\nSourceURL:http://patorjk.com/text-color-fader/\r\n<html>\r\n<body>\r\n<!--StartFragment--><td style=\"color: rgb(255, 0, 0); font-family: &quot;Helvetica Neue&quot;, Helvetica, Arial, sans-serif; font-size: 14px; font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: start; text-indent: 0px; text-transform: none; white-space: pre; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; background-color: rgb(0, 0, 0); text-decoration-style: initial; text-decoration-color: initial; display: inline !important; float: none;\"> </span><!--EndFragment-->\r\n</body>\r\n</html>\0";
            string testColorSmall =
                "Version:0.9\r\nStartHTML:0000000153\r\nEndHTML:0000000767\r\nStartFragment:0000000189\r\nEndFragment:0000000731\r\nSourceURL:http://patorjk.com/text-color-fader/\r\n<html>\r\n<body>\r\n<!--StartFragment--><table><tr><td style=\"color: rgb(255, 0, 0); font-family: &quot;Helvetica Neue&quot;, Helvetica, Arial, sans-serif; font-size: 14px; font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: start; text-indent: 0px; text-transform: none; white-space: pre; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; background-color: rgb(0, 0, 0); text-decoration-style: initial; text-decoration-color: initial; display: inline !important; float: none;\"> </td></tr></table><!--EndFragment-->\r\n</body>\r\n</html>\0";
            string newTestColor =
                "Version:0.9\r\nStartHTML:0000000105\r\nEndHTML:0000000856\r\nStartFragment:0000000141\r\nEndFragment:0000000820\r\n<html>\r\n<body>\r\n<!--StartFragment--><meta name=\"generator\" content=\"Sheets\"/><style type=\"text/css\"><!--td {border: 1px solid #ccc;}br {mso-data-placement:same-cell;}--></style><table xmlns=\"http://www.w3.org/1999/xhtml\" cellspacing=\"0\" cellpadding=\"0\" dir=\"ltr\" border=\"1\" style=\"table-layout:fixed;font-size:10pt;font-family:arial,sans,sans-serif;width:0px;border-collapse:collapse;border:none\"><colgroup><col width=\"65\"/><col width=\"65\"/></colgroup><tbody><tr style=\"height:21px;\"><td style=\"overflow:hidden;padding:2px 3px 2px 3px;vertical-align:bottom;background-color:#00ffff;\"></td><td style=\"overflow:hidden;padding:2px 3px 2px 3px;vertical-align:bottom;background-color:#00ffff;\"></td></tr></tbody></table><!--EndFragment-->\r\n</body>\r\n</html>\0";
            string newTestColorDrai =
                "Version:0.9\r\nStartHTML:0000000105\r\nEndHTML:0000000856\r\nStartFragment:0000000141\r\nEndFragment:0000000820\r\n<html>\r\n<body>\r\n<!--StartFragment--><meta name=\"generator\" content=\"Sheets\"/><style type=\"text/css\"><!--td {border: 1px solid #ccc;}br {mso-data-placement:same-cell;}--></style><table xmlns=\"http://www.w3.org/1999/xhtml\" cellspacing=\"0\" cellpadding=\"0\" dir=\"ltr\" border=\"1\" style=\"table-layout:fixed;font-size:10pt;font-family:arial,sans,sans-serif;width:0px;border-collapse:collapse;border:none\"><colgroup><col width=\"65\"/><col width=\"65\"/></colgroup><tbody><tr style=\"height:21px;\"><td style=\"overflow:hidden;padding:2px 3px 2px 3px;vertical-align:bottom;background-color:#00ffff;\"></td><td style=\"overflow:hidden;padding:2px 3px 2px 3px;vertical-align:bottom;background-color:#00ffff;\"></td><td style=\"overflow:hidden;padding:2px 3px 2px 3px;vertical-align:bottom;background-color:#00ffff;\"></td></tr></tbody></table><!--EndFragment-->\r\n</body>\r\n</html>\0";
            string htmlStart =
                "Version:0.9\r\nStartHTML:aaaaaaaaaa\r\nEndHTML:bbbbbbbbbb\r\nStartFragment:cccccccccc\r\nEndFragment:dddddddddd\r\n";
            string htmlContentStart =
                "<html>\r\n<body>\r\n<!--StartFragment--><meta name=\"generator\" content=\"Sheets\"/><style type=\"text/css\"></style><table xmlns=\"http://www.w3.org/1999/xhtml\" cellspacing=\"0\" cellpadding=\"0\" dir=\"ltr\" style=\"table-layout:fixed;font-size:10pt;font-family:arial,sans,sans-serif;width:0px;border-collapse:collapse;border:none\"><colgroup><col width=\"50\"/><col width=\"50\"/></colgroup><tbody>";


            string uselessString = "<span style =\"color: rgb(0, 0, 0); font-family: &quot;Helvetica Neue&quot;, Helvetica, Arial, sans-serif; font-size: 14px; font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: start; text-indent: 0px; text-transform: none; white-space: pre; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; background-color: rgb(0, 0, 0); text-decoration-style: initial; text-decoration-color: initial; display: inline !important; float: none;\"> </span>";
            string htmlEnd = "</tbody></table><!--EndFragment-->\r\n</body>\r\n</html>\0";
            string rowStart = "<tr style=\"height:21px;\">";
            string rowEnd = "</tr>";
            //Clipboard.SetData(DataFormats.Html, twoColoredHansHans);
            //Clipboard.SetDataObject(new DataObject(DataFormats.Html, firstColor), true);
            var secondColor = Clipboard.GetData(DataFormats.Html);

            var matches = Regex.Matches(newTestColor, hexregex);
            foreach (var match in matches)
            {
                //MessageBox.Show(match.ToString());
            }


            string colorHex = "ff0000";

            var intHexString = Int64.Parse(colorHex, System.Globalization.NumberStyles.HexNumber).ToString();


            var replacesColor = Regex.Replace(newTestColorDrai, decregex, "," + intHexString);
            replacesColor = Regex.Replace(replacesColor, hexregex, "#" + colorHex);
            replacesColor = Regex.Replace(replacesColor, rgbregex, "rgb(50, 205, 50)");
            var replacesColorRGB = Regex.Replace(newTestColor, rgbregex, "rgb(50, 205, 50)");
            //Clipboard.Clear();
            Clipboard.SetData(DataFormats.Html, newTestColor);
            Clipboard.SetDataObject(new DataObject(DataFormats.Html, replacesColor), true);
            Clipboard.SetData(DataFormats.Html, replacesColor);
            Application.Exit();


            //Clipboard.SetData(DataFormats.Html, secondColor);

            HashSet<Color> hash = new HashSet<Color>();
            StringBuilder sb = new StringBuilder();

            var realFilename = filename ?? "WritingPencil.jpg";

            Bitmap img = new Bitmap(realFilename);

            double oldWidth = img.Width;
            double oldHeight = img.Height;
            double newWidth = size;
            double newHeight = newWidth / oldWidth * oldHeight;

            img = ResizeImage(img, (int)newWidth, (int)newHeight);
            for (int hieght = 0; hieght < img.Height; hieght++)
            {
                sb.Append(rowStart);
                for (int width = 0; width < img.Width; width++)
                {
                

                    Color pixel = img.GetPixel(width, hieght);
                    hash.Add(pixel);
                    //var sat = pixel.GetBrightness() - 1;

                    sb.Append(Cell(ColorTranslator.ToHtml(pixel)));

                    //sb.Append("\t");



                }

                sb.Append(rowEnd);
                //sb.Append("\n");
            }

            var sData = htmlStart + htmlContentStart + sb.ToString() + htmlEnd;
            sData = sData.Replace("aaaaaaaaaa", htmlStart.Length.ToString().PadLeft(10, '0'));
            sData = sData.Replace("bbbbbbbbbb", sData.Length.ToString().PadLeft(10, '0'));
            sData = sData.Replace("cccccccccc", (htmlStart + htmlContentStart).Length.ToString().PadLeft(10, '0'));
            sData = sData.Replace("dddddddddd", (htmlStart + htmlContentStart + sb.ToString()).Length.ToString().PadLeft(10, '0'));


            Clipboard.Clear();
            //Clipboard.SetData(DataFormats.Html ,sData);
            Clipboard.SetDataObject(new DataObject(DataFormats.Html, sData), true);
            StringBuilder colorString = new StringBuilder();
            foreach (var color in hash)
            {
                colorString.Append($"{color.GetHashCode()}: {color.ToKnownColor()}\n");
            }
            Application.Exit();
        }


        public static string Cell(string hexcolor)
        {
            var hexregex = @"#[\d|\w]{6}";
            var x =
                "<td style =\"overflow:hidden;padding:2px 3px 2px 3px;vertical-align:bottom;background-color:#00ffff;\"></td>";
            return Regex.Replace(x, hexregex, hexcolor);
        }

        public static Bitmap ResizeImage(Image image, int width, int height)
        {
            var destRect = new Rectangle(0, 0, width, height);
            var destImage = new Bitmap(width, height);

            destImage.SetResolution(image.HorizontalResolution, image.VerticalResolution);

            using (var graphics = Graphics.FromImage(destImage))
            {
                graphics.CompositingMode = CompositingMode.SourceCopy;
                graphics.CompositingQuality = CompositingQuality.HighQuality;
                graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                graphics.SmoothingMode = SmoothingMode.HighQuality;
                graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;

                using (var wrapMode = new ImageAttributes())
                {
                    wrapMode.SetWrapMode(WrapMode.TileFlipXY);
                    graphics.DrawImage(image, destRect, 0, 0, image.Width, image.Height, GraphicsUnit.Pixel, wrapMode);
                }
            }

            return destImage;
        }

        public byte[] GetBytesFromClipboardObj()
        {
            DataObject retrievedData = Clipboard.GetDataObject() as DataObject;
            if (retrievedData == null || !retrievedData.GetDataPresent(typeof(Byte[])))
            {
                MessageBox.Show("Isnull");
                return null;
            }

            return retrievedData.GetData(typeof(Byte[])) as Byte[];
        }


        public void PutBytesOnClipboardObj(byte[] byteArr)
        {
            DataObject data = new DataObject();
            // Can technically just be written as "SetData(byteArr)", but this is more clear.
            data.SetData(typeof(byte[]), byteArr);
            // The 'copy=true' argument means the data will remain available
            // after the program is closed.
            Clipboard.Clear();

            Clipboard.SetDataObject(data, true);

        }

        byte[] ObjectToByteArray(object obj)
        {
            if (obj == null)
                return null;
            BinaryFormatter bf = new BinaryFormatter();
            using (MemoryStream ms = new MemoryStream())
            {
                bf.Serialize(ms, obj);
                return ms.ToArray();
            }
        }

        private void btnChoosePicture_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.ShowDialog();
            var f = ofd.FileName;
            var size = (int)numericUpDown1.Value;
            ClipBoardPicture(f, size);
        }
    }
}

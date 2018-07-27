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
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using WindowsInput;
using WindowsInput.Native;

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

        public void ClipBoardPicture(string filename, int size, bool withMakkro = false)
        {

            Application.Exit();
            if (withMakkro)
            {
                OneStep(filename, size);
                return;
            }
            string htmlStart =
                "Version:0.9\r\nStartHTML:aaaaaaaaaa\r\nEndHTML:bbbbbbbbbb\r\nStartFragment:cccccccccc\r\nEndFragment:dddddddddd\r\n";
            string htmlContentStart =
                "<html>\r\n<body>\r\n<!--StartFragment--><meta name=\"generator\" content=\"Sheets\"/><style type=\"text/css\"></style><table xmlns=\"http://www.w3.org/1999/xhtml\" cellspacing=\"0\" cellpadding=\"0\" dir=\"ltr\" style=\"table-layout:fixed;font-size:10pt;font-family:arial,sans,sans-serif;width:0px;border-collapse:collapse;border:none\"><colgroup><col width=\"50\"/><col width=\"50\"/></colgroup><tbody>";


            string uselessString = "<span style =\"color: rgb(0, 0, 0); font-family: &quot;Helvetica Neue&quot;, Helvetica, Arial, sans-serif; font-size: 14px; font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: start; text-indent: 0px; text-transform: none; white-space: pre; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; background-color: rgb(0, 0, 0); text-decoration-style: initial; text-decoration-color: initial; display: inline !important; float: none;\"> </span>";
            string htmlEnd = "</tbody></table><!--EndFragment-->\r\n</body>\r\n</html>\0";
            string rowStart = "<tr style=\"height:21px;\">";
            string rowEnd = "</tr>";
            StringBuilder sb = new StringBuilder();

 

            Bitmap img = new Bitmap(filename);

            double oldWidth = img.Width;
            double oldHeight = img.Height;
            double newWidth = size;
            double newHeight = newWidth / oldWidth * oldHeight;

            var counter = 0;

            img = ResizeImage(img, (int)newWidth, (int)newHeight);
            for (int hieght = 0; hieght < img.Height; hieght++)
            {
                sb.Append(rowStart);
                for (int width = 0; width < img.Width; width++)
                {
                
                    Color pixel = img.GetPixel(width, hieght);
 
                    sb.Append(Cell(ColorTranslator.ToHtml(pixel)));
                }

                sb.Append(rowEnd);

            }

            var sData = htmlStart + htmlContentStart + sb.ToString() + htmlEnd;
            sData = sData.Replace("aaaaaaaaaa", htmlStart.Length.ToString().PadLeft(10, '0'));
            sData = sData.Replace("bbbbbbbbbb", sData.Length.ToString().PadLeft(10, '0'));
            sData = sData.Replace("cccccccccc", (htmlStart + htmlContentStart).Length.ToString().PadLeft(10, '0'));
            sData = sData.Replace("dddddddddd", (htmlStart + htmlContentStart + sb.ToString()).Length.ToString().PadLeft(10, '0'));


            Clipboard.Clear();
            //Clipboard.SetData(DataFormats.Html ,sData);
            Clipboard.SetDataObject(new DataObject(DataFormats.Html, sData), true);

            Application.Exit();
        }


        public static void OneStep(string filename, int size)
        {
            Thread.Sleep(5000);

            Bitmap img = new Bitmap(filename);

            double oldWidth = img.Width;
            double oldHeight = img.Height;
            double newWidth = size;
            double newHeight = newWidth / oldWidth * oldHeight;

            img = ResizeImage(img, (int)newWidth, (int)newHeight);


            for (int hieght = 0; hieght < img.Height; hieght++)
            {

                //color the row
                for (int width = 0; width < img.Width; width++)
                {

                    Color pixel = img.GetPixel(width, hieght);

                    if (pixel.GetBrightness() < 1)
                    {
                        keyboard.ModifiedKeyStroke(VirtualKeyCode.CONTROL, VirtualKeyCode.VK_V);
                        Thread.Sleep(100);
                    }

                    keyboard.KeyDown(VirtualKeyCode.RIGHT);
                    Thread.Sleep(20);

                }



                for (int width = 0; width < img.Width; width++)
                {
                    keyboard.KeyDown(VirtualKeyCode.LEFT);
                    Thread.Sleep(20);

                }

                keyboard.KeyDown(VirtualKeyCode.DOWN);
                Thread.Sleep(20);
            }


        }



        public static void GreySteps(string filename, int size, int numberOfStepsToDo = 3)
        {
            Thread.Sleep(5000);
   
            Bitmap img = new Bitmap(filename);

            double oldWidth = img.Width;
            double oldHeight = img.Height;
            double newWidth = size;
            double newHeight = newWidth / oldWidth * oldHeight;


            double rangeStep = ((double)1) / numberOfStepsToDo;


            img = ResizeImage(img, (int)newWidth, (int)newHeight);

            //Set flags
            keyboard.KeyDown(VirtualKeyCode.LEFT);
            Thread.Sleep(20);
            Clipboard.SetText(end);
            keyboard.ModifiedKeyStroke(VirtualKeyCode.CONTROL, VirtualKeyCode.VK_V);
            Thread.Sleep(500);
            keyboard.KeyDown(VirtualKeyCode.RIGHT);
            Thread.Sleep(20);

            keyboard.KeyDown(VirtualKeyCode.UP);
            Thread.Sleep(20);
            Clipboard.SetText(end);
            keyboard.ModifiedKeyStroke(VirtualKeyCode.CONTROL, VirtualKeyCode.VK_V);
            Thread.Sleep(500);
            keyboard.KeyDown(VirtualKeyCode.DOWN);
            Thread.Sleep(20);

            //End set flags
            //get color for first row
            keyboard.ModifiedKeyStroke(VirtualKeyCode.SHIFT, VirtualKeyCode.RIGHT);
            Thread.Sleep(50);

            keyboard.ModifiedKeyStroke(VirtualKeyCode.CONTROL, VirtualKeyCode.VK_C);
            Thread.Sleep(50);
            keyboard.KeyDown(VirtualKeyCode.DOWN);
            Thread.Sleep(1000);

            for (int i = 0; i < img.Width; i++)
            {
                keyboard.ModifiedKeyStroke(VirtualKeyCode.CONTROL, VirtualKeyCode.VK_V);
                Thread.Sleep(50);
                keyboard.KeyDown(VirtualKeyCode.RIGHT);
                Thread.Sleep(20);
           
            }

            //set flag for first row
            Clipboard.SetText(end);
            keyboard.ModifiedKeyStroke(VirtualKeyCode.CONTROL, VirtualKeyCode.VK_V);
            Thread.Sleep(50);


            keyboard.KeyDown(VirtualKeyCode.UP);
            Thread.Sleep(20);

            GetBackToStart(false);


            for (int currenStep = 1; currenStep < numberOfStepsToDo; currenStep++)
            {
                var lowRange = currenStep * rangeStep; 
                var highRange = rangeStep * (currenStep + 1);

                for (int hieght = 0; hieght < img.Height; hieght++)
                {
                    //Get Color
                    for (int rightSteps = 0; rightSteps < currenStep; rightSteps++)
                    {
                        keyboard.KeyDown(VirtualKeyCode.RIGHT);
                        Thread.Sleep(20);
                    }
                    keyboard.ModifiedKeyStroke(VirtualKeyCode.CONTROL, VirtualKeyCode.VK_C);
                    Thread.Sleep(200);

                    for (int leftSteps = 0; leftSteps < currenStep; leftSteps++)
                    {
                        keyboard.KeyDown(VirtualKeyCode.LEFT);
                        Thread.Sleep(20);
                    }
                    //end get color
                    
                    //get down for each row
                    keyboard.KeyDown(VirtualKeyCode.DOWN);
                    Thread.Sleep(20);
                    for (int i = 0; i < hieght; i++)
                    {
                        keyboard.KeyDown(VirtualKeyCode.DOWN);
                        Thread.Sleep(20);
                    }

                    //color the row
                    for (int width = 0; width < img.Width; width++)
                    {

                        Color pixel = img.GetPixel(width, hieght);

                        if (pixel.GetBrightness() > lowRange && pixel.GetBrightness() < highRange)
                        {
                            keyboard.ModifiedKeyStroke(VirtualKeyCode.CONTROL, VirtualKeyCode.VK_V);
                            Thread.Sleep(100);
                        }

                        keyboard.KeyDown(VirtualKeyCode.RIGHT);
                        Thread.Sleep(20);
           
                    }

                    GetBackToStart(true);




                    keyboard.KeyDown(VirtualKeyCode.UP);
                    Thread.Sleep(20);
                    GetBackToStart(false);
                }

            }

        }

        private static string end = "kek".GetHashCode().ToString("X");
        private static KeyboardSimulator keyboard = new KeyboardSimulator(new InputSimulator());

        private static void GetBackToStart(bool UpElseLeft = true)
        {
            keyboard.ModifiedKeyStroke(VirtualKeyCode.CONTROL, VirtualKeyCode.VK_C);
            Thread.Sleep(100);
            while (Clipboard.GetText() != end)
            {

                if (UpElseLeft)
                {
                    keyboard.ModifiedKeyStroke(VirtualKeyCode.CONTROL, VirtualKeyCode.UP);

                }
                else
                {
                    keyboard.ModifiedKeyStroke(VirtualKeyCode.CONTROL, VirtualKeyCode.LEFT);
                }
                Thread.Sleep(20);
            }


            if (UpElseLeft)
            {
                keyboard.ModifiedKeyStroke(VirtualKeyCode.CONTROL, VirtualKeyCode.DOWN);

            }
            else
            {
                keyboard.ModifiedKeyStroke(VirtualKeyCode.CONTROL, VirtualKeyCode.RIGHT);
            }
            Thread.Sleep(20);
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

        private void btnChoosePicture_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.ShowDialog();
            var f = ofd.FileName;
            var size = (int)numericUpDown1.Value;
            ClipBoardPicture(f, size, chkWithMakro.Checked);
        }
    }
}

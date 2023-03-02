using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using WIA;
using System.Runtime.InteropServices;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iText.Kernel.Pdf;
using iText.Kernel.Utils;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using System.Web.Helpers;

namespace ScanImage
{
    public partial class Form2 : Form
    {
        Device _device;
        //public const uint FEED_READY = 0x00000001;
        //public const uint FEEDER = 0x00000001;
        //public const uint FLATBED = 0x00000002;

        class WIA_DPS_DOCUMENT_HANDLING_SELECT
        {
            public const int FEEDER = 0x001;
            public const int FLATBED = 0x002;
            public const int DUPLEX = 0x004;
            public const int FRONT_FIRST = 0x008;
            public const int BACK_FIRST = 0x010;
            public const int FRONT_ONLY = 0x020;
            public const int BACK_ONLY = 0x040;
            public const int NEXT_PAGE = 0x080;
            public const int PREFEED = 0x100;
            public const int AUTO_ADVANCE = 0x200;
        }

        class WIA_DPS_DOCUMENT_HANDLING_STATUS
        {
            public const int FEED_READY = 0x01;
            public const int FLAT_READY = 0x02;
            public const int DUP_READY = 0x05;
            public const int FLAT_COVER_UP = 0x08;
            public const int PATH_COVER_UP = 0x10;
            public const int PAPER_JAM = 0x20;
        }

        class WIA_ERRORS
        {
            public const uint BASE_VAL_WIA_ERROR = 0x80210000;
            public const uint WIA_ERROR_GENERAL_ERROR = BASE_VAL_WIA_ERROR + 1;
            public const uint WIA_ERROR_PAPER_JAM = BASE_VAL_WIA_ERROR + 2;
            public const uint WIA_ERROR_PAPER_EMPTY = BASE_VAL_WIA_ERROR + 3;
            public const uint WIA_ERROR_BUSY = BASE_VAL_WIA_ERROR + 6;
        }

        class WIA_PROPERTIES
        {
            private const string WIA_IPS_DOCUMENT_HANDLING_SELECT = "3088";
            public const int WIA_RESERVED_FOR_NEW_PROPS = 1024;
            public const int WIA_DIP_FIRST = 2;
            public const int WIA_DPA_FIRST = WIA_DIP_FIRST + WIA_RESERVED_FOR_NEW_PROPS;
            public const int WIA_DPC_FIRST = WIA_DPA_FIRST + WIA_RESERVED_FOR_NEW_PROPS;
            // Scanner only device properties (DPS)
            public const int WIA_DPS_FIRST = WIA_DPC_FIRST + WIA_RESERVED_FOR_NEW_PROPS;
            public const int WIA_DPS_DOCUMENT_HANDLING_STATUS = WIA_DPS_FIRST + 13;
            public const int WIA_DPS_DOCUMENT_HANDLING_SELECT = WIA_DPS_FIRST + 14;
            public const int WIA_DPS_PAGES = WIA_DPS_FIRST + 22;
        }

        public Form2()
        {
            InitializeComponent();
            //Get default device Id  
            var _deviceId = FindDefaultDeviceId();
            //Find Device  
            var _deviceInfo = FindDevice(_deviceId);
            //Connect the device  
            _device = _deviceInfo.Connect();
        }

        private DeviceInfo FindDevice(string deviceId)
        {
            DeviceManager manager = new DeviceManager();
            foreach (DeviceInfo info in manager.DeviceInfos)
                if (info.DeviceID == deviceId)
                    return info;
            return null;
        }

        private string FindDefaultDeviceId()
        {
            WIA.ICommonDialog dialog = new WIA.CommonDialog();
            WIA.Device device = dialog.ShowSelectDevice(WIA.WiaDeviceType.UnspecifiedDeviceType, false, false);
            //string deviceId = Properties.Settings.Default.ScannerDeviceID; // "{6BDD1FC6-810F-11D0-BEC7-08002BE2092F}\0000";
            string deviceId = device.DeviceID;
            if (String.IsNullOrEmpty(deviceId))
            {
                // Select a scanner  
                WIA.CommonDialog wiaDiag = new WIA.CommonDialog();
                Device d = wiaDiag.ShowSelectDevice(WiaDeviceType.ScannerDeviceType, true, false);
                if (d != null)
                {
                    deviceId = d.DeviceID;
                    Properties.Settings.Default.ScannerDeviceID = deviceId;
                    Properties.Settings.Default.Save();
                }
            }
            return deviceId;
        }

        //public List<System.Drawing.Image> ScanPages(int dpi = 300, double width = 8.5, double height = 14, bool isSingleSide)
        public List<System.Drawing.Image> ScanPages(int dpi, double width, double height, bool isSingleSide)
        {
            //WIA.ICommonDialog dialog = new WIA.CommonDialog();
            //_device = dialog.ShowSelectDevice(WIA.WiaDeviceType.UnspecifiedDeviceType, true, false);           
            Item item = _device.Items[1];
            // configure item of the device              
            SetDeviceItemProperty(ref item, 6146, 1); // greyscale  
            SetDeviceItemProperty(ref item, 6147, dpi); // 150 dpi  
            SetDeviceItemProperty(ref item, 6148, dpi); // 150 dpi  
            SetDeviceItemProperty(ref item, 6151, (int)(dpi * width)); // set scan width  
            SetDeviceItemProperty(ref item, 6152, (int)(dpi * height)); // set scan height  
            //SetDeviceItemProperty(ref item, 4104, 8); // bit depth  
            // Detect if the ADF is loaded, if not use the flatbed  
            List<System.Drawing.Image> images = GetPagesFromScanner(ScanSourceSingleSode.DocumentFeeder, ScanSourceDoubleSide.DocumentFeeder, isSingleSide, item);
            if (images.Count == 0)
            {
                // check the flatbed if ADF is not loaded, try from flatbed  
                DialogResult dialogResult;
                do
                {
                    //List<Image> singlePage = GetPagesFromScanner(ScanSource.Flatbed, item);
                    List<System.Drawing.Image> singlePage = GetPagesFromScanner(ScanSourceSingleSode.DocumentFeeder, ScanSourceDoubleSide.DocumentFeeder, isSingleSide, item);
                    images.AddRange(singlePage);
                    dialogResult = MessageBox.Show("Do you want to scan another page?", "ScanToEvernote", MessageBoxButtons.YesNo);
                }
                while (dialogResult == DialogResult.Yes);
            }
            return images;
        }

        private List<System.Drawing.Image> GetPagesFromScanner(ScanSourceSingleSode singlesource, ScanSourceDoubleSide doublesource, bool isSingleSide, Item item)
        {
            SetDeviceProperty(ref _device, 3096, 1);
            bool callingFrom = false;
            List<System.Drawing.Image> images = new List<System.Drawing.Image>();
            int handlingStatus = GetDeviceProperty(ref _device, WIA_PROPERTIES.WIA_DPS_DOCUMENT_HANDLING_STATUS);
            if (isSingleSide)
            {
                SetDeviceProperty(ref _device, 3088, (int)ScanSourceSingleSode.DocumentFeeder);
                if ((singlesource == ScanSourceSingleSode.DocumentFeeder && handlingStatus == WIA_DPS_DOCUMENT_HANDLING_STATUS.FEED_READY) || (singlesource == ScanSourceSingleSode.Flatbed && handlingStatus == WIA_DPS_DOCUMENT_HANDLING_STATUS.FLAT_READY) || (singlesource == ScanSourceSingleSode.DocumentFeeder && handlingStatus == WIA_DPS_DOCUMENT_HANDLING_STATUS.DUP_READY))
                {
                    callingFrom = true;
                }
            }
            else
            {
                SetDeviceProperty(ref _device, 3088, (int)ScanSourceDoubleSide.DocumentFeeder);
                if ((doublesource == ScanSourceDoubleSide.DocumentFeeder && handlingStatus == WIA_DPS_DOCUMENT_HANDLING_STATUS.FEED_READY) || (doublesource == ScanSourceDoubleSide.Flatbed && handlingStatus == WIA_DPS_DOCUMENT_HANDLING_STATUS.FLAT_READY) || (doublesource == ScanSourceDoubleSide.DocumentFeeder && handlingStatus == WIA_DPS_DOCUMENT_HANDLING_STATUS.DUP_READY))
                {
                    callingFrom = true;
                }
            }


            //if ((source == ScanSource.DocumentFeeder && handlingStatus == WIA_DPS_DOCUMENT_HANDLING_STATUS.FEED_READY) || (source == ScanSource.Flatbed && handlingStatus == WIA_DPS_DOCUMENT_HANDLING_STATUS.FLAT_READY) || (source == ScanSource.DocumentFeeder && handlingStatus == WIA_DPS_DOCUMENT_HANDLING_STATUS.DUP_READY)) //FLATBED_READY))
            if (callingFrom) //FLATBED_READY))
            {
                do
                {
                    ImageFile wiaImage = null;
                    try
                    {

                        wiaImage = (ImageFile)item.Transfer(FormatID.wiaFormatJPEG); // WIA_FORMAT_JPEG



                        //ImageFile image = (ImageFile)item.Transfer(FormatID.wiaFormatJPEG);

                        //// Save all images to file
                        //int i = 0;
                        //do
                        //{
                        //    // Save the image
                        //    string fileName = string.Format(@"c:\Users\Administrator\Desktop\scan_{0}.jpg", i);
                        //    image.SaveFile(fileName);

                        //    // Move to the next image (if any)
                        //    i++;
                        //    try
                        //    {
                        //        image = (ImageFile)item.Transfer(FormatID.wiaFormatJPEG);
                        //    }
                        //    catch
                        //    {
                        //        // No more images
                        //        image = null;
                        //    }
                        //} while (image != null);

                    }
                    catch (COMException ex)
                    {
                        if ((uint)ex.ErrorCode == WIA_ERRORS.WIA_ERROR_PAPER_EMPTY)
                            break;
                        else
                            throw;
                    }
                    if (wiaImage != null)
                    {
                        System.Diagnostics.Trace.WriteLine(String.Format("Image is {0} x {1} pixels", (float)wiaImage.Width / 150, (float)wiaImage.Height / 150));
                        System.Drawing.Image image = ConvertToImage(wiaImage);

                        //string fileName = Path.GetTempFileName();
                        //File.Delete(fileName);
                        //wiaImage.SaveFile(fileName);
                        //wiaImage = null;
                        // add file to output list
                        //images.Add(System.Drawing.Image.FromFile(fileName));
                        images.Add(image);
                    }
                }
                while (true);
                //while (source == ScanSource.DocumentFeeder);
            }
            return images;
        }

        private static System.Drawing.Image ConvertToImage(ImageFile wiaImage)
        {
             byte[] imageBytes = (byte[])wiaImage.FileData.get_BinaryData();
             long quality = 30;
            //WebImage img = new WebImage(imageBytes);
            //if (img.Width > 1000)
            //    //img.Resize(1000, 1000);
            //    img.Resize(800, 800);
            //imageBytes = img.GetBytes();

             using (var inputStream = new MemoryStream(imageBytes))
             using (var outputStream = new MemoryStream())
             {
                 using (var image1 = System.Drawing.Image.FromStream(inputStream))
                 {
                     var encoder = GetEncoderInfo(ImageFormat.Jpeg);
                     var encoderParams = new EncoderParameters(2);
                     encoderParams.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, quality);

                     // Set additional compression settings
                     var qualityParam = new EncoderParameter(System.Drawing.Imaging.Encoder.Compression, (long)EncoderValue.CompressionCCITT4);
                     encoderParams.Param[1] = qualityParam;

                     image1.Save(outputStream, encoder, encoderParams);
                     imageBytes = outputStream.ToArray();
                 }
             }

            MemoryStream ms = new MemoryStream(imageBytes);
            System.Drawing.Image image = System.Drawing.Image.FromStream(ms);
            return image;
        }

        private static ImageCodecInfo GetEncoderInfo(ImageFormat format)
        {
            return ImageCodecInfo.GetImageEncoders().FirstOrDefault(codec => codec.FormatID == format.Guid);
        }

        #region Get / set device properties
        private void SetDeviceProperty(ref Device device, int propertyID, int propertyValue)
        {
            foreach (Property p in device.Properties)
            {
                if (p.PropertyID == propertyID)
                {
                    object value = propertyValue;
                    p.set_Value(ref value);
                    break;
                }
            }
        }

        private int GetDeviceProperty(ref Device device, int propertyID)
        {
            int ret = -1;
            foreach (Property p in device.Properties)
            {
                if (p.PropertyID == 3088)
                {
                    ret = (int)p.get_Value();
                    break;
                }
            }
            return ret;
        }

        private void SetDeviceItemProperty(ref Item item, int propertyID, int propertyValue)
        {
            foreach (Property p in item.Properties)
            {
                if (p.PropertyID == propertyID)
                {
                    object value = propertyValue;
                    p.set_Value(ref value);
                    break;
                }
            }
        }

        private int GetDeviceItemProperty(ref Item item, int propertyID)
        {
            int ret = -1;
            foreach (Property p in item.Properties)
            {
                if (p.PropertyID == propertyID)
                {
                    ret = (int)p.get_Value();
                    break;
                }
            }
            return ret;
        }
        #endregion

        private void button1_Click(object sender, EventArgs e)
        {
            bool isSingleSide = false;

            if (checkBox2.Checked)
            {
                isSingleSide = true;
            }
            else if (checkBox1.Checked)
            {
                isSingleSide = false;
            }
            //CompressPdf(@"C:\Users\Administrator\Desktop\scanDOUBLE.pdf", @"C:\Users\Administrator\Desktop\comp.pdf");

            ////int dpi = 150; double width = 8.5; double height = 14; //for legal Paper Size
            int dpi = 150; double width = 8.35; double height = 11.70; //For A4 Paper Size
            List<System.Drawing.Image> obj = ScanPages(dpi, width, height, isSingleSide);

            byte[] document = ImageToByteArray(obj);
            ////byte[] document = ConvertIntoSinglePDF(obj);
            File.WriteAllBytes(@"C:\Users\Administrator\Desktop\scan.pdf", document);

            //This code i written to upload the file  
            //int a = 0;
            //foreach (System.Drawing.Image aa in obj)
            //{
            //    a++;
            //    //aa.Save("D:\\ScanerUploadedFileWIA\\myfile.png _" + DateTime.Now.ToLongTimeString(), ImageFormat.Png);  
            //    aa.Save(@"C:\Users\Administrator\Desktop\" + a + ".jpeg", ImageFormat.Jpeg);
            //}
        }

        enum ScanSourceSingleSode
        {
            DocumentFeeder = 1,
            Flatbed = 2,
        }

        enum ScanSourceDoubleSide
        {
            DocumentFeeder = 5,
            Flatbed = 2,
        }

        public byte[] ImageToByteArray(List<System.Drawing.Image> imageIn)
        {
            //var ms = new MemoryStream();
            Document imageDocument = null;
            iTextSharp.text.pdf.PdfWriter imageDocumentWriter = null;
            Document doc = new Document();
            doc.SetPageSize(PageSize.A4);
            var ms1 = new System.IO.MemoryStream();
            {
                PdfCopy pdf = new PdfCopy(doc, ms1);
                doc.Open();
                for (int i = 0; i < imageIn.Count; i++)
                {
                    //using (var ms = new MemoryStream())
                    var ms = new MemoryStream();
                    //{
                    imageIn[i].Save(ms, imageIn[i].RawFormat);
                    //}
                    byte[] data = ms.ToArray();
                    imageDocument = new Document();
                    using (var imageMS = new MemoryStream())
                    {
                        imageDocumentWriter = iTextSharp.text.pdf.PdfWriter.GetInstance(imageDocument, imageMS);
                        imageDocument.Open();
                        if (imageDocument.NewPage())
                        {
                            var image = iTextSharp.text.Image.GetInstance(data);
                            image.Alignment = Element.ALIGN_CENTER;
                            //image.ScaleToFit(doc.PageSize.Width - 10, doc.PageSize.Height - 10);
                            image.ScaleToFit(doc.PageSize.Width - 7, doc.PageSize.Height - 45);
                            if (!imageDocument.Add(image))
                            {
                                throw new Exception("Unable to add image to page!");
                            }
                            imageDocument.Close();
                            imageDocumentWriter.Close();
                            iTextSharp.text.pdf.PdfReader imageDocumentReader = new iTextSharp.text.pdf.PdfReader(imageMS.ToArray());
                            var page = pdf.GetImportedPage(imageDocumentReader, 1);
                            pdf.AddPage(page);
                            imageDocumentReader.Close();
                        }
                    }
                }
            }
            if (doc.IsOpen()) doc.Close();
            return ms1.ToArray();
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }

        public static void CompressPdf(string inputFilePath, string outputFilePath)
        {

            using (var inputFileStream = new FileStream(inputFilePath, FileMode.Open))
            using (var outputFileStream = new FileStream(outputFilePath, FileMode.Create))
            {
                var writer = new iText.Kernel.Pdf.PdfWriter(outputFileStream, new WriterProperties().SetCompressionLevel(CompressionConstants.BEST_COMPRESSION));
                var pdfDoc = new iText.Kernel.Pdf.PdfDocument(new iText.Kernel.Pdf.PdfReader(inputFileStream), writer);

                pdfDoc.Close();
            }

            //using (var inputFileStream = new FileStream(inputFilePath, FileMode.Open))
            //using (var outputFileStream = new FileStream(outputFilePath, FileMode.Create))
            //{
            //    var writer = new iText.Kernel.Pdf.PdfWriter(outputFileStream);
            //    var reader = new iText.Kernel.Pdf.PdfReader(inputFileStream);
            //    var pdfDocument = new iText.Kernel.Pdf.PdfDocument(reader, writer);

            //    var compressor = new PdfCompressor();
            //    compressor.SetCompressionLevel(9);
            //    compressor.SetRemoveUnusedObjects(true);
            //    compressor.SetEmbedFonts(true);

            //    var result = new iText.Kernel.Pdf.PdfDocument(new PdfWriter(new MemoryStream()));
            //    compressor.Compress(pdfDocument, result);

            //    pdfDocument.Close();
            //    result.Close();
            //}

        }

        //public static byte[] ConvertIntoSinglePDF(List<string> filePaths)
        //{
        //    Document doc = new Document();
        //    doc.SetPageSize(PageSize.A4);

        //    var ms = new System.IO.MemoryStream();
        //    {
        //        PdfCopy pdf = new PdfCopy(doc, ms);
        //        doc.Open();

        //        foreach (string path in filePaths)
        //        {
        //            byte[] data = File.ReadAllBytes(path);
        //            doc.NewPage();
        //            Document imageDocument = null;
        //            PdfWriter imageDocumentWriter = null;
        //            switch (Path.GetExtension(path).ToLower().Trim('.'))
        //            {
        //                case "bmp":
        //                case "gif":
        //                case "jpg":
        //                case "png":
        //                    imageDocument = new Document();
        //                    using (var imageMS = new MemoryStream())
        //                    {
        //                        imageDocumentWriter = PdfWriter.GetInstance(imageDocument, imageMS);
        //                        imageDocument.Open();
        //                        if (imageDocument.NewPage())
        //                        {
        //                            var image = iTextSharp.text.Image.GetInstance(data);
        //                            image.Alignment = Element.ALIGN_CENTER;
        //                            image.ScaleToFit(doc.PageSize.Width - 10, doc.PageSize.Height - 10);
        //                            if (!imageDocument.Add(image))
        //                            {
        //                                throw new Exception("Unable to add image to page!");
        //                            }
        //                            imageDocument.Close();
        //                            imageDocumentWriter.Close();
        //                            PdfReader imageDocumentReader = new PdfReader(imageMS.ToArray());
        //                            var page = pdf.GetImportedPage(imageDocumentReader, 1);
        //                            pdf.AddPage(page);
        //                            imageDocumentReader.Close();
        //                        }
        //                    }
        //                    break;
        //                case "pdf":
        //                    var reader = new PdfReader(data);
        //                    for (int i = 0; i < reader.NumberOfPages; i++)
        //                    {
        //                        pdf.AddPage(pdf.GetImportedPage(reader, i + 1));
        //                    }
        //                    pdf.FreeReader(reader);
        //                    reader.Close();
        //                    break;
        //                default:
        //                    break;
        //            }
        //        }

        //        if (doc.IsOpen()) doc.Close();
        //        return ms.ToArray();
        //    }
        //}
    }
}


using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using WIA;

namespace ScanImage
{
    public partial class Form3 : Form
    {
        public Form3()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //WIA.DeviceManager manager = new WIA.DeviceManager();

            //// Find the first scanner device in the list
            //WIA.DeviceInfo scannerInfo = null;
            //foreach (WIA.DeviceInfo info in manager.DeviceInfos)
            //{
            //    if (info.Type == WIA.WiaDeviceType.ScannerDeviceType)
            //    {
            //        scannerInfo = info;
            //        break;
            //    }
            //}

            //if (scannerInfo == null)
            //{
            //    throw new Exception("Scanner not found");
            //}

            //// Connect to the scanner device
            //WIA.Device scannerDevice = scannerInfo.Connect();

            //// Set the duplex scanning mode to scan both sides of the document
            //foreach (WIA.Property property in scannerDevice.Properties)
            //{
            //    if (property.PropertyID == 3088)
            //    {
            //        property.set_Value(1);
            //        break;
            //    }
            //}

            //// Create a new item for the scanned image
            //WIA.Item item = scannerDevice.Items.Add(WIA.WiaItemType.File);

            //// Set the image format to JPEG
            //WIA.Property imageFormat = item.Properties["WIA_IPA_FORMAT"];
            //imageFormat.set_Value(FormatID.wiaFormatJPEG);

            //// Start the scan and retrieve the scanned image(s)
            //WIA.ICommonDialog commonDialog = new WIA.CommonDialog();
            //WIA.Item frontItem = null;
            //WIA.Item backItem = null;
            //try
            //{
            //    frontItem = commonDialog.ShowAcquireImage(WIA.WiaDeviceType.ScannerDeviceType, WIA.WiaImageIntent.ColorIntent, WIA.WiaImageBias.MaximizeQuality, WIA_IPS_CUR_INTENT.getCurIntent(), item, false, true, false);
            //    backItem = commonDialog.ShowAcquireImage(WIA.WiaDeviceType.ScannerDeviceType, WIA.WiaImageIntent.ColorIntent, WIA.WiaImageBias.MaximizeQuality, WIA_IPS_CUR_INTENT.getCurIntent(), item, false, true, false);
            //}
            //catch (Exception ex)
            //{
            //    // Handle scan errors
            //}

            //// Save the scanned images to disk
            //frontItem.SaveFile("front.jpg");
            //backItem.SaveFile("back.jpg");



        }
    }
}



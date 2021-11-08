using AtomicSeller.Models;
using AtomicSeller.ViewModels;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security.Authentication;
using System.Text;
using System.Threading.Tasks;
using System.Web.Mvc;
using System.Web.Script.Serialization;
using ShopifyOrderAPI.Models;
using iTextSharp.text;
using iTextSharp.text.pdf;
using Microsoft.Win32;
using System.Text.RegularExpressions;
using AtomicSeller.Helpers;

namespace AtomicSeller
{
    class HeppnerEDI
    {

        public void MakeIFCSUMFile(List<DeliveryDM> DonneeEDIFACTList, string PDFDirPath, DBSchenkerSettings _DBSchenkerSettings) // DonneeEDIFACTList est une liste avec toutes les Deliveries
        {

            DEALVM _DEALVM = new DEALVM();

            string filepath = PDFDirPath + "EDIIFCSUMD96AFile.txt";

            //DeliveryDM _data = DonneeEDIFACTList[0];
            new IFCSUMD96A().IFCSUMD96A_AddHeaderToFile(DonneeEDIFACTList, filepath, _DBSchenkerSettings);

            new IFCSUMD96A().IFCSUMD96A_AddDEALVMSenderToFile(DonneeEDIFACTList, filepath, _DBSchenkerSettings);

            //IFCSUMD96A_AddDEALVMReceiverToFile(_DeliveryDM);
            new IFCSUMD96A().IFCSUMD96A_AddFooterToFile(DonneeEDIFACTList, filepath, _DBSchenkerSettings);

        }


        public void MakeInternationalLabels(List<DeliveryDM> _Deliveries, string PDFDirPath, DBSchenkerSettings _DBSchenkerSettings)
        {
            int deliveryIndex = 0;
            foreach (var _Delivery in _Deliveries)
            {
                deliveryIndex += 1;
                string PDFFilePath = PDFDirPath + "InternationalLabel" + deliveryIndex.ToString() + ".pdf";


                string Weight = _Delivery.ParcelWeight;
                int width = 104;
                int height = 111;
                // Make PDF label
                using (FileStream fs = new FileStream(PDFFilePath, FileMode.Create, FileAccess.Write, FileShare.None))
                {
                    #region Page
                    double PageWidth = iTextSharp.text.Utilities.MillimetersToPoints(width);
                    double PageHeight = iTextSharp.text.Utilities.MillimetersToPoints(height);

                    double PageLeftMargin = 0;
                    double PageRightMargin = 0;
                    double PageTopMargin = 0;
                    double PageBottomMargin = 0;
                    #endregion


                    string agenceGTF = "219013"; // GTF of DBSchenker BLOIS, this data is constant.
                    string codeRegime = "3"; // Regime code of SystemFrance(01) is 3 (last page of CDC) 

                    //platforms:
                    List<string> platforms = new List<string>();


                    Document document = new Document(new iTextSharp.text.Rectangle(
                        (float)PageWidth,
                        (float)PageHeight),
                        (float)PageLeftMargin,
                        (float)PageRightMargin,
                        (float)PageTopMargin,
                        (float)PageBottomMargin);

                    PdfWriter writer = PdfWriter.GetInstance(document, fs);

                    // Add meta information to the document
                    document.AddCreator("AtomicSeller");
                    document.Open();

                    // Makes it possible to add text to a specific place in the document using 
                    // a X & Y placement syntax.
                    PdfContentByte cb = writer.DirectContent;
                    cb.BeginText();

                    Font F111_300 = AtomicGetFont("Calibri", 7f, "Normal");
                    Font F211_300 = AtomicGetFont("Calibri", 9f, "Bold");
                    Font F232_300 = AtomicGetFont("Calibri", 9f, "Normal");
                    Font F311_300 = AtomicGetFont("Calibri", 11f, "Bold");
                    Font F411_300 = AtomicGetFont("Calibri", 12f, "Bold");
                    Font F511_300 = AtomicGetFont("Calibri", 19f, "Bold");
                    Font F222_300 = AtomicGetFont("Calibri", 16f, "Normal");
                    Font F312_300 = AtomicGetFont("Calibri", 13f, "Bold");
                    Font F322_300 = AtomicGetFont("Calibri", 22f, "Normal");
                    Font F323_300 = AtomicGetFont("Calibri", 22f, "Bold");
                    Font F334_300 = AtomicGetFont("Calibri", 20f, "Bold");
                    Font F623_300_1 = AtomicGetFont("Calibri", 30f, "Bold", 1);
                    /********************************************************************************/
                    //OutLines((2, 1309) - (1226, 2))
                    //Top
                    float startX56 = (float)2;
                    float startX57 = (float)1226;
                    float startY56 = (float)1305;
                    float startY57 = (float)1309;
                    DrawRectangleInPoints(startX56, startY56, startX57, startY57, cb, 1);

                    //Right
                    float startX58 = (float)1222;
                    float startX59 = (float)1226;
                    float startY58 = (float)2;
                    float startY59 = (float)1309;
                    DrawRectangleInPoints(startX58, startY58, startX59, startY59, cb, 1);

                    //Bottom
                    float startX60 = (float)2;
                    float startX61 = (float)1226;
                    float startY60 = (float)2;
                    float startY61 = (float)6;
                    DrawRectangleInPoints(startX60, startY60, startX61, startY61, cb, 1);

                    //Left
                    float startX54 = (float)2;
                    float startX55 = (float)6;
                    float startY54 = (float)2;
                    float startY55 = (float)1309;
                    DrawRectangleInPoints(startX54, startY54, startX55, startY55, cb, 1);

                    //Header
                    string InternationalLogoPath = "Resources/img/InternationalLogo.png";
                    string ImgPath = Path.Combine(System.AppDomain.CurrentDomain.BaseDirectory, InternationalLogoPath);
                    float startX75 = (float)25;
                    float scaleX = (float)350;
                    float startY75 = (float)1210;
                    float scaleY = (float)80;
                    InsertImage(ImgPath, cb, startX75, startY75, scaleX, scaleY);

                    float startX2 = (float)370;
                    float startY2 = (float)1270;
                    WriteRotateText_Dot(cb, startX2, startY2, "0", F232_300, "carrier");

                    float startX1 = (float)25;
                    float startY1 = (float)1210;
                    WriteRotateText_Dot(cb, startX1, startY1, "0", F232_300, "HEPPNER GONESSE Tel.+33130116666 fax. +33130116666");

                    float startX5 = (float)10;
                    float startX6 = (float)1218;
                    float startY5 = (float)1155;
                    float startY6 = (float)1153;
                    DrawRectangleInPoints(startX5, startY5, startX6, startY6, cb);

                    //Sender
                    float startX9 = (float)25;
                    float startY9 = (float)1143;
                    WriteRotateText_Dot(cb, startX9, startY9, "0", F232_300, "From (exp.) : ");

                    float startX10 = (float)250;
                    float startY10 = (float)1143;
                    WriteRotateText_Dot(cb, startX10, startY10, "0", F232_300, _Delivery.SenderZip);

                    float startX11 = (float)25;
                    float startY11 = (float)1093;
                    WriteRotateText_Dot(cb, startX11, startY11, "0", F211_300, "Sender name");

                    float startX12 = (float)25;
                    float startY12 = (float)1053;
                    WriteRotateText_Dot(cb, startX12, startY12, "0", F232_300, _Delivery.SenderName);

                    float startX7 = (float)10;
                    float startX8 = (float)1218;
                    float startY7 = (float)945;
                    float startY8 = (float)943;
                    DrawRectangleInPoints(startX7, startY7, startX8, startY8, cb);

                    //Recipient
                    float startX15 = (float)606;
                    float startY15 = (float)1143;
                    WriteRotateText_Dot(cb, startX15, startY15, "0", F232_300, "To (Dest.) : ");

                    float startX16 = (float)831;
                    float startY16 = (float)1143;
                    WriteRotateText_Dot(cb, startX16, startY16, "0", F232_300, _Delivery.RecipZip);

                    float startX17 = (float)616;
                    float startY17 = (float)1093;
                    WriteRotateText_Dot(cb, startX17, startY17, "0", F211_300, "Recipient Name");

                    float startX18 = (float)616;
                    float startY18 = (float)1053;
                    WriteRotateText_Dot(cb, startX18, startY18, "0", F232_300, _Delivery.RecipCompanyName);

                    float startX19 = (float)616;
                    float startY19 = (float)1013;
                    WriteRotateText_Dot(cb, startX19, startY19, "0", F232_300, "FROM 26180 RASTEDE");

                    float startX13 = (float)600;
                    float startX14 = (float)602;
                    float startY13 = (float)1153;
                    float startY14 = (float)945;
                    DrawRectangleInPoints(startX13, startY13, startX14, startY14, cb);

                    //Shipment
                    float startX22 = (float)25;
                    float startY22 = (float)935;
                    WriteRotateText_Dot(cb, startX22, startY22, "0", F232_300, "Receipt:  : ");

                    float startX23 = (float)250;
                    float startY23 = (float)935;
                    WriteRotateText_Dot(cb, startX23, startY23, "0", F232_300, _Delivery.RecipZip);

                    float startX24 = (float)25;
                    float startY24 = (float)885;
                    WriteRotateText_Dot(cb, startX24, startY24, "0", F232_300, "Shipping date : ");

                    float startX25 = (float)270;
                    float startY25 = (float)885;
                    WriteRotateText_Dot(cb, startX25, startY25, "0", F211_300, _Delivery.ShippingDate.ToString("yy/MM/dd"));

                    float startX26 = (float)450;
                    float startY26 = (float)885;
                    WriteRotateText_Dot(cb, startX26, startY26, "0", F232_300, "Delivery date : ");

                    float startX27 = (float)675;
                    float startY27 = (float)885;
                    WriteRotateText_Dot(cb, startX27, startY27, "0", F211_300, _Delivery.DepositDate.ToString("yy/MM/dd"));

                    float startX28 = (float)25;
                    float startY28 = (float)835;
                    WriteRotateText_Dot(cb, startX28, startY28, "0", F232_300, "Instruction : ");

                    float startX29 = (float)250;
                    float startY29 = (float)835;
                    WriteRotateText_Dot(cb, startX29, startY29, "0", F211_300, "Make an appointment at the 0033654654");

                    float startX30 = (float)250;
                    float startY30 = (float)805;
                    WriteRotateText_Dot(cb, startX30, startY30, "0", F211_300, "Tailgate needed");

                    float startX20 = (float)10;
                    float startX21 = (float)1218;
                    float startY20 = (float)745;
                    float startY21 = (float)743;
                    DrawRectangleInPoints(startX20, startY20, startX21, startY21, cb);

                    //Country && Transport

                    float startX31 = (float)10;
                    float startX32 = (float)1218;
                    float startY31 = (float)375;
                    float startY32 = (float)373;
                    DrawRectangleInPoints(startX31, startY31, startX32, startY32, cb);


                    float startX35 = (float)10;
                    float startX36 = (float)410;
                    float startY35 = (float)745;
                    float startY36 = (float)373;
                    DrawRectangleInPoints(startX35, startY35, startX36, startY36, cb);


                    float startX37 = (float)818;
                    float startX38 = (float)1218;
                    float startY37 = (float)745;
                    float startY38 = (float)560;
                    DrawRectangleInPoints(startX37, startY37, startX38, startY38, cb);

                    float startX81 = (float)818;
                    float startX82 = (float)1218;
                    float startY81 = (float)548;
                    float startY82 = (float)373;
                    DrawRectangleInPoints(startX81, startY81, startX82, startY82, cb);


                    float startX51 = (float)40;
                    float startY51 = (float)725;
                    WriteRotateText_Dot(cb, startX51, startY51, "0", F623_300_1, "FROM");

                    float startX52 = (float)40;
                    float startY52 = (float)615;
                    WriteRotateText_Dot(cb, startX52, startY52, "0", F623_300_1, "26180");



                    float startX53 = (float)898;
                    float startY53 = (float)730;
                    WriteRotateText_Dot(cb, startX53, startY53, "0", F623_300_1, "I108");

                    float startX80 = (float)908;
                    float startY80 = (float)570;
                    WriteRotateText_Dot(cb, startX80, startY80, "0", F623_300_1, "000");

                    float startX39 = (float)420;
                    float startY39 = (float)735;
                    WriteRotateText_Dot(cb, startX39, startY39, "0", F232_300, "Weight : ");

                    float startX40 = (float)555;
                    float startY40 = (float)735;
                    WriteRotateText_Dot(cb, startX40, startY40, "0", F211_300, _Delivery.ParcelWeight + " kg");

                    float startX41 = (float)420;
                    float startY41 = (float)685;
                    WriteRotateText_Dot(cb, startX41, startY41, "0", F232_300, "Parcel : ");

                    float startX42 = (float)555;
                    float startY42 = (float)685;
                    WriteRotateText_Dot(cb, startX42, startY42, "0", F211_300, _Delivery.ParcelLenght);

                    float startX43 = (float)420;
                    float startY43 = (float)635;
                    WriteRotateText_Dot(cb, startX43, startY43, "0", F232_300, "Port of : ");

                    float startX44 = (float)555;
                    float startY44 = (float)635;
                    WriteRotateText_Dot(cb, startX44, startY44, "0", F211_300, "NO");

                    float startX45 = (float)420;
                    float startY45 = (float)585;
                    WriteRotateText_Dot(cb, startX45, startY45, "0", F232_300, "C.R : ");

                    float startX46 = (float)555;
                    float startY46 = (float)585;
                    WriteRotateText_Dot(cb, startX46, startY46, "0", F211_300, "NO");

                    float startX47 = (float)420;
                    float startY47 = (float)535;
                    WriteRotateText_Dot(cb, startX47, startY47, "0", F232_300, "V.D. : ");

                    float startX48 = (float)555;
                    float startY48 = (float)535;
                    WriteRotateText_Dot(cb, startX48, startY48, "0", F211_300, "NO");

                    //Parcel
                    float startX49 = (float)25;
                    float startY49 = (float)363;
                    WriteRotateText_Dot(cb, startX49, startY49, "0", F232_300, "Reference : ");

                    float startX50 = (float)250;
                    float startY50 = (float)363;
                    WriteRotateText_Dot(cb, startX50, startY50, "0", F211_300, _Delivery.senderParcelRef);

                    float startX33 = (float)10;
                    float startX34 = (float)1218;
                    float startY33 = (float)285;
                    float startY34 = (float)283;
                    DrawRectangleInPoints(startX33, startY33, startX34, startY34, cb);

                    //Barcode
                    string Barcode = "*" + agenceGTF + _DBSchenkerSettings.ChargerCode + _Delivery.UniqueIdentifier + codeRegime + "00" + "'";

                    float startX72 = (float)400;
                    float startX73 = (float)1000;
                    float startY72 = (float)265;
                    float startY73 = (float)100;
                    InsertBarCode128_mm(Barcode, cb, startX72, startY72, startX73, startY73);

                    float startX74 = (float)500;
                    float startY74 = (float)100;
                    WriteRotateText_Dot(cb, startX74, startY74, "0", F111_300, Barcode);
                    /********************************************************************************/

                    cb.EndText();

                    document.Close();
                    writer.Close();
                    fs.Close();

                }
            }
        }


        public void MakeNationalLabels(List<DeliveryDM> _Deliveries, string PDFDirPath, DBSchenkerSettings _DBSchenkerSettings)
        {
            int deliveryIndex = 0;
            foreach (var _Delivery in _Deliveries)
            {
                deliveryIndex += 1;
                string PDFFilePath = PDFDirPath + "NationalLabel" + deliveryIndex.ToString() + ".pdf";


                string Weight = _Delivery.ParcelWeight;
                int width = 104;
                int height = 111;
                // Make PDF label
                using (FileStream fs = new FileStream(PDFFilePath, FileMode.Create, FileAccess.Write, FileShare.None))
                {
                    #region Page
                    double PageWidth = iTextSharp.text.Utilities.MillimetersToPoints(width);
                    double PageHeight = iTextSharp.text.Utilities.MillimetersToPoints(height);

                    double PageLeftMargin = 0;
                    double PageRightMargin = 0;
                    double PageTopMargin = 0;
                    double PageBottomMargin = 0;
                    #endregion


                    string agenceGTF = "219013"; // GTF of DBSchenker BLOIS, this data is constant.
                    string codeRegime = "3"; // Regime code of SystemFrance(01) is 3 (last page of CDC) 

                    //platforms:
                    List<string> platforms = new List<string>();


                    Document document = new Document(new iTextSharp.text.Rectangle(
                        (float)PageWidth,
                        (float)PageHeight),
                        (float)PageLeftMargin,
                        (float)PageRightMargin,
                        (float)PageTopMargin,
                        (float)PageBottomMargin);

                    PdfWriter writer = PdfWriter.GetInstance(document, fs);

                    // Add meta information to the document
                    document.AddCreator("AtomicSeller");
                    document.Open();

                    // Makes it possible to add text to a specific place in the document using 
                    // a X & Y placement syntax.
                    PdfContentByte cb = writer.DirectContent;
                    cb.BeginText();

                    Font F111_300 = AtomicGetFont("Calibri", 7f, "Normal");
                    Font F211_300 = AtomicGetFont("Calibri", 9f, "Bold");
                    Font F311_300 = AtomicGetFont("Calibri", 11f, "Bold");
                    Font F411_300 = AtomicGetFont("Calibri", 12f, "Bold");
                    Font F511_300 = AtomicGetFont("Calibri", 19f, "Bold");
                    Font F222_300 = AtomicGetFont("Calibri", 16f, "Normal");
                    Font F312_300 = AtomicGetFont("Calibri", 13f, "Bold");
                    Font F322_300 = AtomicGetFont("Calibri", 22f, "Normal");
                    Font F323_300 = AtomicGetFont("Calibri", 22f, "Bold");
                    Font F323_300_1 = AtomicGetFont("Calibri", 22f, "Bold", 1);
                    Font F623_300_1 = AtomicGetFont("Calibri", 60f, "Bold", 1);
                    Font F334_300 = AtomicGetFont("Calibri", 20f, "Bold");
                    /********************************************************************************/

                    //OutLines((2, 1309) - (1226, 2))
                    //Top
                    float startX56 = (float)2;
                    float startX57 = (float)1226;
                    float startY56 = (float)1305;
                    float startY57 = (float)1309;
                    DrawRectangleInPoints(startX56, startY56, startX57, startY57, cb, 1);

                    //Right
                    float startX58 = (float)1222;
                    float startX59 = (float)1226;
                    float startY58 = (float)2;
                    float startY59 = (float)1309;
                    DrawRectangleInPoints(startX58, startY58, startX59, startY59, cb, 1);

                    //Bottom
                    float startX60 = (float)2;
                    float startX61 = (float)1226;
                    float startY60 = (float)2;
                    float startY61 = (float)6;
                    DrawRectangleInPoints(startX60, startY60, startX61, startY61, cb, 1);

                    //Left
                    float startX54 = (float)2;
                    float startX55 = (float)6;
                    float startY54 = (float)2;
                    float startY55 = (float)1309;
                    DrawRectangleInPoints(startX54, startY54, startX55, startY55, cb, 1);

                    //Content Part
                    //Sender Part

                    float startX1 = (float)25;
                    float startY1 = (float)1300;
                    WriteRotateText_Dot(cb, startX1, startY1, "0", F411_300, "Exp : ");

                    float startX2 = (float)150;
                    float startY2 = (float)1300;
                    WriteRotateText_Dot(cb, startX2, startY2, "0", F411_300, _Delivery.SenderCompanyName.ToUpper());

                    float startX3 = (float)1000;
                    float startY3 = (float)1300;
                    WriteRotateText_Dot(cb, startX3, startY3, "0", F311_300, _Delivery.SenderZip);

                    float startX5 = (float)10;
                    float startX6 = (float)1218;
                    float startY5 = (float)1205;
                    float startY6 = (float)1203;
                    DrawRectangleInPoints(startX5, startY5, startX6, startY6, cb);

                    //Reciptient
                    float startX7 = (float)25;
                    float startY7 = (float)1195;
                    WriteRotateText_Dot(cb, startX7, startY7, "0", F411_300, "Dest : ");

                    float startX8 = (float)250;
                    float startY8 = (float)1195;
                    WriteRotateText_Dot(cb, startX8, startY8, "0", F411_300, _Delivery.RecipCompanyName.ToUpper());

                    float startX9 = (float)250;
                    float startY9 = (float)1135;
                    WriteRotateText_Dot(cb, startX9, startY9, "0", F211_300, _Delivery.RecipFirstName.ToUpper());

                    float startX10 = (float)250;
                    float startY10 = (float)1090;
                    WriteRotateText_Dot(cb, startX10, startY10, "0", F211_300, _Delivery.RecipAdr1.ToUpper());

                    float startX11 = (float)250;
                    float startY11 = (float)1055;
                    WriteRotateText_Dot(cb, startX11, startY11, "0", F411_300, _Delivery.RecipCountryISOCode.ToUpper());

                    float startX34 = (float)10;
                    float startX35 = (float)1218;
                    float startY34 = (float)985;
                    float startY35 = (float)983;
                    DrawRectangleInPoints(startX34, startY34, startX35, startY35, cb);

                    //Shipment
                    float startX12 = (float)25;
                    float startY12 = (float)975;
                    WriteRotateText_Dot(cb, startX12, startY12, "0", F411_300, "Recepisse n. :");

                    float startX14 = (float)350;
                    float startY14 = (float)975;
                    WriteRotateText_Dot(cb, startX14, startY14, "0", F411_300, _Delivery.senderParcelRef);

                    float startX13 = (float)25;
                    float startY13 = (float)925;
                    WriteRotateText_Dot(cb, startX13, startY13, "0", F411_300, "Parcel n. : ");

                    float startX15 = (float)350;
                    float startY15 = (float)925;
                    WriteRotateText_Dot(cb, startX15, startY15, "0", F411_300, _Delivery.RecipINSEECode);

                    float startX16 = (float)10;
                    float startX17 = (float)650;
                    float startY16 = (float)845;
                    float startY17 = (float)843;
                    DrawRectangleInPoints(startX16, startY16, startX17, startY17, cb);

                    //Date
                    float startX18 = (float)25;
                    float startY18 = (float)835;
                    WriteRotateText_Dot(cb, startX18, startY18, "0", F411_300, "Date Exp  :");
                    float startX22 = (float)350;
                    float startY22 = (float)835;
                    WriteRotateText_Dot(cb, startX22, startY22, "0", F411_300, _Delivery.ShippingDate.ToString("yy/MM/dd"));

                    float startX19 = (float)25;
                    float startY19 = (float)785;
                    WriteRotateText_Dot(cb, startX19, startY19, "0", F411_300, "Date Liv :");
                    float startX23 = (float)350;
                    float startY23 = (float)785;
                    WriteRotateText_Dot(cb, startX23, startY23, "0", F411_300, _Delivery.DepositDate.ToString("yy/MM/dd"));

                    float startX20 = (float)25;
                    float startY20 = (float)735;
                    WriteRotateText_Dot(cb, startX20, startY20, "0", F411_300, "Weight :");
                    float startX24 = (float)350;
                    float startY24 = (float)735;
                    WriteRotateText_Dot(cb, startX24, startY24, "0", F411_300, _Delivery.ParcelWeight);

                    float startX21 = (float)25;
                    float startY21 = (float)685;
                    WriteRotateText_Dot(cb, startX21, startY21, "0", F411_300, "C.R. : ");
                    float startX25 = (float)350;
                    float startY25 = (float)685;
                    WriteRotateText_Dot(cb, startX25, startY25, "0", F411_300, "NO");

                    float startX43 = (float)10;
                    float startX44 = (float)650;
                    float startY43 = (float)615;
                    float startY44 = (float)613;
                    DrawRectangleInPoints(startX43, startY43, startX44, startY44, cb);

                    //Transport
                    float startX26 = (float)25;
                    float startY26 = (float)605;
                    WriteRotateText_Dot(cb, startX26, startY26, "0", F411_300, "Bay : ");

                    float startX49 = (float)350;
                    float startX50 = (float)640;
                    float startY49 = (float)493;
                    float startY50 = (float)605;
                    DrawRectangleInPoints(startX49, startY49, startX50, startY50, cb);
                    float startX51 = (float)420;
                    float startY51 = (float)615;
                    WriteRotateText_Dot(cb, startX51, startY51, "0", F323_300_1, "015");


                    float startX52 = (float)650;
                    float startX53 = (float)1218;
                    float startY52 = (float)985;
                    float startY53 = (float)485;
                    DrawRectangleInPoints(startX52, startY52, startX53, startY53, cb);
                    float startX27 = (float)800;
                    float startY27 = (float)985;
                    WriteRotateText_Dot(cb, startX27, startY27, "0", F623_300_1, "67");
                    float startX28 = (float)680;
                    float startY28 = (float)685;
                    WriteRotateText_Dot(cb, startX28, startY28, "0", F323_300_1, "STRASBOURG");

                    float startX45 = (float)10;
                    float startX46 = (float)1218;
                    float startY45 = (float)485;
                    float startY46 = (float)483;
                    DrawRectangleInPoints(startX45, startY45, startX46, startY46, cb);

                    float startX47 = (float)648;
                    float startX48 = (float)650;
                    float startY47 = (float)485;
                    float startY48 = (float)985;
                    DrawRectangleInPoints(startX47, startY47, startX48, startY48, cb);

                    //BarCode
                    float startX29 = (float)10;
                    float startX30 = (float)1218;
                    float startY29 = (float)100;
                    float startY30 = (float)98;
                    DrawRectangleInPoints(startX29, startY29, startX30, startY30, cb);

                    string Barcode = "*" + agenceGTF + _DBSchenkerSettings.ChargerCode + _Delivery.UniqueIdentifier + codeRegime + "00" + "'";

                    float startX72 = (float)150;
                    float startX73 = (float)1078;
                    float startY72 = (float)455;
                    float startY73 = (float)355;
                    InsertBarCode128_mm(Barcode, cb, startX72, startY72, startX73, startY73);

                    float startX74 = (float)370;
                    float startY74 = (float)320;
                    WriteRotateText_Dot(cb, startX74, startY74, "0", F111_300, Barcode);

                    //Footer
                    string InternationalLogoPath = "Resources/img/NationalLogo.png"; 
                    string ImgPath = Path.Combine(System.AppDomain.CurrentDomain.BaseDirectory, InternationalLogoPath);
                    float startX75 = (float)25;
                    float scaleX = (float)310;
                    float startY75 = (float)10;
                    float scaleY = (float)80;
                    InsertImage(ImgPath, cb, startX75, startY75, scaleX, scaleY);

                    float start31 = (float)350;
                    float startY31 = (float)60;
                    WriteRotateText_Dot(cb, start31, startY31, "0", F211_300, "Heppner GONESSE - Tel 01 49 39 39 39");

                    float startX32 = (float)1000;
                    float startX33 = (float)998;
                    float startY32 = (float)96;
                    float startY33 = (float)4;
                    DrawRectangleInPoints(startX32, startY32, startX33, startY33, cb);

                    float startX71 = (float)1050;
                    float startY71 = (float)100;
                    WriteRotateText_Dot(cb, startX71, startY71, "0", F411_300, "Mess");

                    /********************************************************************************/

                    cb.EndText();

                    document.Close();
                    writer.Close();
                    fs.Close();

                }
            }
        }


        public void DrawRectangleInPoints(float llx_p, float lly_p, float urx_p, float ury_p, PdfContentByte cb, int BGColorID = 0)
        {
            //change dots to points
            llx_p = iTextSharp.text.Utilities.MillimetersToPoints(iTextSharp.text.Utilities.InchesToMillimeters(llx_p / 300));
            lly_p = iTextSharp.text.Utilities.MillimetersToPoints(iTextSharp.text.Utilities.InchesToMillimeters(lly_p / 300));
            urx_p = iTextSharp.text.Utilities.MillimetersToPoints(iTextSharp.text.Utilities.InchesToMillimeters(urx_p / 300));
            ury_p = iTextSharp.text.Utilities.MillimetersToPoints(iTextSharp.text.Utilities.InchesToMillimeters(ury_p / 300));

            cb.EndText();
            var rect = new iTextSharp.text.Rectangle(llx_p, lly_p, urx_p, ury_p);
            //rect.Border = iTextSharp.text.Rectangle.LEFT_BORDER | iTextSharp.text.Rectangle.RIGHT_BORDER;
            rect.Border = iTextSharp.text.Rectangle.BOX;
            rect.BorderWidth = 1;
            if (BGColorID == 0)
            {
                rect.BackgroundColor = new BaseColor(0, 0, 0);
                rect.BorderColor = new BaseColor(0, 0, 0);
            }
            else
            {
                rect.BackgroundColor = new BaseColor(255, 0, 0);
                rect.BorderColor = new BaseColor(255, 0, 0);
            }
            cb.Rectangle(rect);

            cb.Stroke();
            cb.BeginText();

        }

        public void InsertImage(string ImagePath, PdfContentByte cb, float Leftx, float Lowery, float Scalex, float Scaley)
        {
            float LeftX_p = iTextSharp.text.Utilities.MillimetersToPoints(iTextSharp.text.Utilities.InchesToMillimeters(Leftx / 300));// 200;
            float LowY_p = iTextSharp.text.Utilities.MillimetersToPoints(iTextSharp.text.Utilities.InchesToMillimeters(Lowery / 300));

            float _scaleX = iTextSharp.text.Utilities.MillimetersToPoints(iTextSharp.text.Utilities.InchesToMillimeters((Scalex) / 300));
            float _scaleY = iTextSharp.text.Utilities.MillimetersToPoints(iTextSharp.text.Utilities.InchesToMillimeters(Scaley / 300));
            using (Stream inputImageStream = new FileStream(ImagePath, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                iTextSharp.text.Image sigimage = iTextSharp.text.Image.GetInstance(inputImageStream);
                sigimage.SetAbsolutePosition(LeftX_p, LowY_p);
                sigimage.ScaleAbsolute(_scaleX, _scaleY);
                cb.AddImage(sigimage);
            }
        }
        public void InsertBarCode128_mm(string Code, PdfContentByte cb, float Leftx, float Lowery, float Rightx, float Uppery)
        {
            //float LeftX_p = iTextSharp.text.Utilities.MillimetersToPoints(Leftx_mm);
            //float LowY_p = iTextSharp.text.Utilities.MillimetersToPoints(Lowery_mm);
            //float RightX_p = iTextSharp.text.Utilities.MillimetersToPoints(Rightx_mm);
            //float UpY_p = iTextSharp.text.Utilities.MillimetersToPoints(Uppery_mm);
            //change dots to points
            float LeftX_p = iTextSharp.text.Utilities.MillimetersToPoints(iTextSharp.text.Utilities.InchesToMillimeters(Leftx / 300));// 200;
            float LowY_p = iTextSharp.text.Utilities.MillimetersToPoints(iTextSharp.text.Utilities.InchesToMillimeters(Lowery / 300));
            float RightX_p = iTextSharp.text.Utilities.MillimetersToPoints(iTextSharp.text.Utilities.InchesToMillimeters((Rightx) / 300));
            float UpY_p = iTextSharp.text.Utilities.MillimetersToPoints(iTextSharp.text.Utilities.InchesToMillimeters(Uppery / 300));

            Rectangle BarCodeRect = new Rectangle(LeftX_p, LowY_p, RightX_p, UpY_p);

            Barcode128 ItextSharpBarCode = new Barcode128
            {
                Code = Code,
                TextAlignment = Element.ALIGN_CENTER,
                StartStopText = false,//true,
                CodeType = Barcode.CODE128,
                Extended = false,
                Font = null
            };
            Image _img = ItextSharpBarCode.CreateImageWithBarcode(cb, BaseColor.BLACK, BaseColor.BLACK);
            _img.ScaleAbsolute(BarCodeRect);
            _img.SetAbsolutePosition(LeftX_p, LowY_p);
            cb.AddImage(_img);
        }

        public void WriteRotateText_Dot(PdfContentByte cb, float StartX, float StartY, string Rotation, Font _Font, string Text)
        {
            //chage dots to points
            StartX = iTextSharp.text.Utilities.MillimetersToPoints(iTextSharp.text.Utilities.InchesToMillimeters(StartX / 300));
            StartY = iTextSharp.text.Utilities.MillimetersToPoints(iTextSharp.text.Utilities.InchesToMillimeters(StartY / 300));


            cb.EndText();

            PdfPCell cell = new PdfPCell();
            PdfPTable table = new PdfPTable(1);
            table.TotalWidth = 400;
            //Paragraph paragraph = new Paragraph(Text);
            Phrase Phrase1 = new Phrase(new Chunk(Text, _Font));
            //paragraph.Font = _Font;

            //StartX -= 500;
            //StartY /= 2;

            cell = new PdfPCell(Phrase1);
            if (Rotation == "1")
            {
                cell.Rotation = 270;
                StartX = StartX - 395;
            }
            if (Rotation == "2")
            {
                cell.Rotation = 180;
                StartX = StartX - 395;
            }
            if (Rotation == "3")
            {
                cell.Rotation = 90;
                StartX = StartX - 395;
            }

            cell.BorderWidth = 0;

            table.AddCell(cell);
            table.WriteSelectedRows(0, -1, StartX, StartY, cb);
            cb.BeginText();

        }

        public iTextSharp.text.Font AtomicGetFont(string fontName, float FontSize, string FontStyle, int ColorType = 0)
        {

            //// FontSize
            if (FontSize < 1) FontSize = 2;

            //// FontStyle
            bool Bold = false, Italic = false;

            int FontStyleToken = iTextSharp.text.Font.NORMAL;
            if (FontStyle.ToUpper().Contains("BOLD")) { FontStyleToken = iTextSharp.text.Font.BOLD; Bold = true; }
            if (FontStyle.ToUpper().Contains("ITALIC")) { FontStyleToken = FontStyleToken | iTextSharp.text.Font.ITALIC; Italic = true; }
            if (FontStyle.ToUpper().Contains("UNDERLINE")) FontStyleToken = FontStyleToken | iTextSharp.text.Font.UNDERLINE;

            //// FontName
            string filename;
            if (string.IsNullOrEmpty(fontName))
                filename = "calibri.ttf";
            else
            {
                //System.Drawing.Font TestFont = new System.Drawing.Font( );
                filename = GetSystemFontFileName(fontName, Bold, Italic);
                //filename = fontName + ".ttf";
                //Local.AtomicMainForm.Text += filename;
            }

            if (!FontFactory.IsRegistered(fontName))
            {
                var fontsFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.Fonts);

                var fontPath = Environment.GetEnvironmentVariable("SystemRoot") + "\\fonts\\" + filename;
                FontFactory.Register(fontPath);
            }

            if (ColorType == 0)
            {
                return FontFactory.GetFont(fontName, BaseFont.IDENTITY_H, BaseFont.EMBEDDED, FontSize, FontStyleToken);
            }
            else if (ColorType == 1)
            {
                return FontFactory.GetFont(fontName, BaseFont.IDENTITY_H, BaseFont.EMBEDDED, FontSize, FontStyleToken, new BaseColor(255, 255, 255));
            }
            else
            {
                return FontFactory.GetFont(fontName, BaseFont.IDENTITY_H, BaseFont.EMBEDDED, FontSize, FontStyleToken);
            }
        }

        public string GetSystemFontFileName(string fontname, bool Bold, bool Italic) //System.Drawing.Font font)
        {
            RegistryKey fonts = null;

            try
            {
                fonts = Registry.LocalMachine.OpenSubKey(@"Software\Microsoft\Windows NT\CurrentVersion\Fonts", false);
                if (fonts == null)
                {
                    fonts = Registry.LocalMachine.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Fonts", false);
                    if (fonts == null)
                    {
                        throw new Exception("Can't find font registry database.");
                    }
                }

                string suffix = "";
                if (Bold)
                    suffix += "(?: Bold)?";
                if (Italic)
                    suffix += "(?: Italic)?";

                var regex = new Regex(@"^(?:.+ & )?" + Regex.Escape(fontname) + @"(?: & .+)?(?<suffix>" + suffix + @") \(TrueType\)$", RegexOptions.Compiled);

                string[] names = fonts.GetValueNames();

                string name = names.Select(n => regex.Match(n)).Where(m => m.Success).OrderByDescending(m => m.Groups["suffix"].Length).Select(m => m.Value).FirstOrDefault();

                if (name != null)
                {
                    return fonts.GetValue(name).ToString();
                }
                else
                {
                    return null;
                }
            }
            finally
            {
                if (fonts != null)
                {
                    fonts.Dispose();
                }
            }
        }


    }
}
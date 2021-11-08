using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Globalization;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using AtomicSeller.Helpers;
using AtomicSeller.ViewModels;
using System.Net.Http;
using Newtonsoft.Json.Linq;
using AtomicSeller;
using System.IO;

namespace AtomicSeller.Controllers
{
    public class HomeController : BaseController
    {
        [HttpGet]
        public ActionResult Index()
        {
            //SessionBag.SetSessionBagInitData();
            return View();
        }

        DBSchenkerSettings InitDBSchenkerSettings()
        {
            DBSchenkerSettings _DBSchenkerSettings = new DBSchenkerSettings();

            /*
            Agence DB Schenker 41 BLOIS :
            Téléphone : 02.54.56.33.40 
            SIRET : 31179945600737 
            Code GTF : 219013 
            Code 3LC : XBQ 
            Code BELT : 41
             */

            _DBSchenkerSettings.MessageId = "0000275";
            _DBSchenkerSettings.AccountNumber = "3255"; // Billing Account of VF Source
            _DBSchenkerSettings.UMNumber = "1";
            _DBSchenkerSettings.AgenceDBSchenkerName = "SCHENKER FRANCE";
            _DBSchenkerSettings.AgenceDBSchenkerCity = "BLOIS";
            _DBSchenkerSettings.AgenceDBSchenkerPhone = "TEL: 02-54-56-33-40";

            // Bar Code Data Init
            _DBSchenkerSettings.RecipINSEECode = "85";
            _DBSchenkerSettings.DirectionalCodeGTF = "851";
            _DBSchenkerSettings.AgencyGTF = "219013";
            _DBSchenkerSettings.ChargerCode = "420000";
            _DBSchenkerSettings.RegimeCode = "3";

            _DBSchenkerSettings.SenderSIRET = "83367052400010";
            _DBSchenkerSettings.SenderSIREN = "311799456"; 
            _DBSchenkerSettings.Reference = "TEST REF";

            _DBSchenkerSettings.DeliveryMode = "T";
            _DBSchenkerSettings.DeliveryTour = "30";

            return _DBSchenkerSettings;
        }

        List<DeliveryDM> InitiDeliveries()
        {
            List<DeliveryDM> Deliveries = new List<ViewModels.DeliveryDM>();

            AtomicSeller.ViewModels.DeliveryDM _Delivery1 = new ViewModels.DeliveryDM();

            _Delivery1.senderParcelRef = "42006157"; // RECEPISSE !!!!
            _Delivery1.RecipINSEECode = "21410"; //INSEE est optionnel; si on connait pas de INSEE on peut utilisr aussi aussi la code postale

            _Delivery1.UniqueIdentifier = _Delivery1.senderParcelRef + "001"; // Unique Identifier (11 carac) == recipisse (8 carac) + order number (3 carac)
            _Delivery1.ShippingSupport = "PE";
            _Delivery1.Platforms = "MONT";

            _Delivery1.Shipment_ID = "42006157001";
            _Delivery1.ProductCode = "0001";
            _Delivery1.ParcelShipmentSequence = 1;
            _Delivery1.DeliveryInstructions = "Livrer au dépôt B";
            _Delivery1.RecipFirstName = "Mr Jean Pierrer Papin";
            _Delivery1.ParcelWeight = "10.0";
            _Delivery1.parcelNumberPartner = "*85851" + "31179945600737";
            _Delivery1.SenderCompanyName = "VF Solutions";
            _Delivery1.SenderAddr1 = "16 Rue du Trieu Quesnoy";
            _Delivery1.SenderAddr2 = "";
            _Delivery1.SenderZip = "59115";
            _Delivery1.SenderCity = "Lyon";
            _Delivery1.SenderName = "STOCK LOGISTIC";
            _Delivery1.SenderPhoneNumber = "0320200903";
            _Delivery1.SenderEmail = "vfsols@atomicseller.fr";
            _Delivery1.SendercountryCode = "FR";

            _Delivery1.RecipCompanyName = "GEMMA";
            _Delivery1.RecipAdr1 = "Rue Henri Dunant";
            _Delivery1.RecipAdr2 = "Etage 3";
            _Delivery1.RecipAdr3 = "Batiment 4";
            _Delivery1.RecipZip = "21510";
            _Delivery1.RecipCity = "MEULSON";
            _Delivery1.RecipCountryLib = "France";
            _Delivery1.RecipLastName = "GLASS";
            _Delivery1.RecipPhoneNumber = "+3378855299339";
            _Delivery1.RecipEmail = "test@atomicseller.com";
            _Delivery1.RecipCountryISOCode = "FR";

            _Delivery1.ShippingDate = new DateTime(2021, 02, 24);
            _Delivery1.DepositDate = new DateTime(2021, 07, 05);
            _Delivery1.TrackingNumber = "127743821";

            Deliveries.Add(_Delivery1);

            AtomicSeller.ViewModels.DeliveryDM _Delivery2 = new ViewModels.DeliveryDM(); ;

            _Delivery2.senderParcelRef = "42006158";

            _Delivery2.UniqueIdentifier = _Delivery2.senderParcelRef + "002";
            _Delivery2.Platforms = "LYON";
            _Delivery2.Shipment_ID = "42006157002"; 
            _Delivery2.ParcelShipmentSequence = 1;
            _Delivery2.MerchandiseNature = "Composants électronique";
            _Delivery2.DeliveryInstructions = "Ne pas laisser sans surveillance";
             
            _Delivery2.RecipFirstName = "Mr Zinedine Zidane";
            _Delivery2.ParcelWeight = "15.0";
            _Delivery2.ParcelValue = "1.0";
            _Delivery2.parcelNumberPartner = "*85851" + "31179945600737";

            _Delivery2.SenderCompanyName = "Client Agence";
            _Delivery2.SenderAddr1 = "5 Allée de la Garonne";
            _Delivery2.SenderAddr2 = "";
            _Delivery2.SenderZip = "99202";
            _Delivery2.SenderCity = "Versailles";
            _Delivery2.SenderName = "STOCK LOGISTIC";
            _Delivery2.SenderPhoneNumber = "0320200903";
            _Delivery2.SenderEmail = "but@atomicseller.fr";
            _Delivery2.SendercountryCode = "FR";

            _Delivery2.RecipCompanyName = "BOHome";
            _Delivery2.RecipAdr1 = "10 Rue Herni Matisse";
            _Delivery2.RecipAdr2 = "Etage 130";
            _Delivery2.RecipAdr3 = "Batiment C";
            _Delivery2.RecipZip = "33520";
            _Delivery2.RecipCity = "BRUGES";
            _Delivery2.RecipCountryLib = "France";
            _Delivery2.RecipLastName = "Grosjean";
            _Delivery2.RecipPhoneNumber = "+332222222222";
            _Delivery2.RecipEmail = "Home@atomicseller.com";
            _Delivery2.RecipCountryISOCode = "FR";

            _Delivery2.ShippingDate = new DateTime(2021, 02, 24);
            _Delivery2.DepositDate = new DateTime(2021, 06, 22);
            _Delivery2.TrackingNumber = "127743821";

            Deliveries.Add(_Delivery2);

            return Deliveries;
        }

        [HttpGet]
        public ActionResult Label()
        {

            List<DeliveryDM> Deliveries = InitiDeliveries();

            DBSchenkerSettings _DBSchenkerSettings = InitDBSchenkerSettings();

            string PDFDirPath = @"D:/AppData/";

            new HeppnerEDI().MakeInternationalLabels(Deliveries, PDFDirPath, _DBSchenkerSettings);
            new HeppnerEDI().MakeNationalLabels(Deliveries, PDFDirPath, _DBSchenkerSettings);

            return RedirectToAction("Index", "Home");
        }

        [HttpGet]
        public ActionResult File()
        {
            List<DeliveryDM> Deliveries = InitiDeliveries();

            DBSchenkerSettings _DBSchenkerSettings = InitDBSchenkerSettings();


            string PDFDirPath = @"D:/AppData/";

            new HeppnerEDI().MakeIFCSUMFile(Deliveries, PDFDirPath, _DBSchenkerSettings);

            return RedirectToAction("Index", "Home");
        }



    }
}
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Web;
using AtomicSeller;
using AtomicSeller.ViewModels;

namespace AtomicSeller
{
    public class IFCSUMD96A
    {
        public void IFCSUMD96A_AddHeaderToFile(List<DeliveryDM> _DeliveryDM, string filepath, DBSchenkerSettings _DBSchenkerSettings)
        {

            if (File.Exists(filepath)) File.Delete(filepath);

            IFCSUMD96A_AddDEALVMHeaderToFile(_DeliveryDM, filepath, _DBSchenkerSettings);
            
        }

        public void IFCSUMD96A_AddDEALVMHeaderToFile(List<DeliveryDM> _DeliveryDM, string filePath, DBSchenkerSettings _DBSchenkerSettings)
        {
            // Add header

            UNA _UNA = new IFCSUMD96A.UNA();
            _UNA.Segment = "UNA";
            _UNA.Record = ":+.? '";
            AppendObjectToEDIFile(_UNA, typeof(UNA), filePath, "");


            // Ici le Code SIREN VF Solutions : VF Solutions (813575024)
            string sirenvfsols = "813575024";

            UNB _UNB = new IFCSUMD96A.UNB();
            _UNB.Segment = "UNB"; 
            _UNB.Record = "+UNOC:1+" + sirenvfsols + ":22 +SCHENKERRW:ZZ+160720:1340+99999'";
            AppendObjectToEDIFile(_UNB, typeof(UNB), filePath, "");

            UNH _UNH = new IFCSUMD96A.UNH();
            _UNH.Segment = "UNH";
            _UNH.Record = "+1+IFCSUM:D:96A:UN'";
            AppendObjectToEDIFile(_UNH, typeof(UNH), filePath, "");
        }

        public void IFCSUMD96A_AddDEALVMSenderToFile(List<DeliveryDM> _DeliveryDMs, string filePath, DBSchenkerSettings _DBSchenkerSettings)
        {
            BGM _BGM = new IFCSUMD96A.BGM();
            _BGM.Segment = "BGM";
            _BGM.Record = "+610+" + _DBSchenkerSettings.MessageId + "'";
            AppendObjectToEDIFile(_BGM, typeof(BGM), filePath, "");

            // Compte de facturation de VF Solutions
            string billingaccountVF = "4122";

            RFF _RFF = new IFCSUMD96A.RFF();
            _RFF.Segment = "RFF";
            _RFF.Record = "+ADE:" + (billingaccountVF + _DeliveryDMs[0].ProductCode) + "'"; 
            AppendObjectToEDIFile(_RFF, typeof(RFF), filePath, "");

            TDT _TDT = new IFCSUMD96A.TDT();
            _TDT.Segment = "TDT";
            _TDT.Record = "+20'";
            AppendObjectToEDIFile(_TDT, typeof(TDT), filePath, "");


            // Ici le Code SIRET Expéditeur : VF Solutions (83367052400010)
            // sender = expediteur
            //_DBSchenkerSettings.SenderSIRET = "83367052400010"; 

            NAD _NAD1 = new IFCSUMD96A.NAD();
            _NAD1.Segment = "NAD";
            _NAD1.Record = "+CZ+" + _DBSchenkerSettings.SenderSIRET + "++" + _DeliveryDMs[0].SenderCompanyName
                                            + ":" + _DeliveryDMs[0].SenderName
                                            + ":" + _DeliveryDMs[0].SenderEmail
                                            + "+" + _DeliveryDMs[0].SenderAddr1
                                            + ":" + _DeliveryDMs[0].SenderAddr2
                                            + ":" + _DeliveryDMs[0].SenderAddr3
                                            + "+" + _DeliveryDMs[0].SenderCity
                                            + "++" + _DeliveryDMs[0].SenderZip
                                            + "+" + _DeliveryDMs[0].SendercountryCode + "'";

            AppendObjectToEDIFile(_NAD1, typeof(NAD), filePath, "");


            // Ici le Code SIRET d'agence d'enlevement : DB Schenker (31179945600737)
            NAD _NAD2 = new IFCSUMD96A.NAD();
            _NAD2.Segment = "NAD";
            
            _NAD2.Record = "+CA+31179945600737++"+ _DBSchenkerSettings.AgenceDBSchenkerName + "'";
            AppendObjectToEDIFile(_NAD2, typeof(NAD), filePath, "");

            //------------------------------------------------------------------//
            // Important :
            // A partir d'Ici on aura des boucles qu'on peut repeter pour autant de Deliveries qu'on a.
            //------------------------------------------------------------------//
             
            foreach (DeliveryDM _Delivery in _DeliveryDMs)
            {
                CNI _CNI = new IFCSUMD96A.CNI();
                _CNI.Segment = "CNI";
                _CNI.Record = "+" + _Delivery.ParcelShipmentSequence + "+" + _Delivery.senderParcelRef + "'";
                AppendObjectToEDIFile(_CNI, typeof(CNI), filePath, "");

                string datetext = _Delivery.DepositDate.ToString("dd/MM/yyyy").Replace("/","");
                DTM _DTM = new IFCSUMD96A.DTM();
                _DTM.Segment = "DTM";
                _DTM.Record = "+2:" + datetext + ":102'"; // 102 pour Date
                AppendObjectToEDIFile(_DTM, typeof(DTM), filePath, "");

                // Il faut ajouter CONDITIONS pour tous les 3 types de FTX:
                if (_Delivery.MerchandiseNature != null)
                {
                    FTX _FTX1 = new IFCSUMD96A.FTX();
                    _FTX1.Segment = "FTX";
                    _FTX1.Record = "+AAA+++" + _Delivery.MerchandiseNature + "'";
                    AppendObjectToEDIFile(_FTX1, typeof(FTX), filePath, "");
                }

                if (_Delivery.DeliveryInstructions != null)
                {
                    FTX _FTX2 = new IFCSUMD96A.FTX();
                    _FTX2.Segment = "FTX";
                    _FTX2.Record = "+DEL+++" + _Delivery.DeliveryInstructions + "'";
                    AppendObjectToEDIFile(_FTX2, typeof(FTX), filePath, "");
                }

                if (_Delivery.DeliveryRemarks != null)
                {
                    FTX _FTX3 = new IFCSUMD96A.FTX();
                    _FTX3.Segment = "FTX";
                    _FTX3.Record = "+DEL+++" + _Delivery.DeliveryInstructions + "'";
                    AppendObjectToEDIFile(_FTX3, typeof(FTX), filePath, "");
                }


                TOD _TOD = new IFCSUMD96A.TOD();
                _TOD.Segment = "TOD";
                _TOD.Record = "+6+P" + "'";
                AppendObjectToEDIFile(_TOD, typeof(TOD), filePath, "");

                RFF _RFF1 = new IFCSUMD96A.RFF();
                _RFF1.Segment = "RFF";
                _RFF1.Record = "+CU:" + _Delivery.RecipCompanyName + "'";
                AppendObjectToEDIFile(_RFF1, typeof(RFF), filePath, "");


                // NAD+CZ et NAD+CA sont unique pour tous les deliveries dans la liste
                // NAD+CN n'est pas unique et peut varier donc on met dans le boucle
                // informations de Destinataire : donc pas de SIRET !
                NAD _NAD3 = new IFCSUMD96A.NAD();
                _NAD3.Segment = "NAD";
                _NAD3.Record = "+CN++" + _Delivery.RecipCompanyName
                                       + ":" + _Delivery.RecipFirstName
                                       + ":" + _Delivery.RecipLastName
                                       + "+" + _Delivery.RecipAdr1
                                       + ":" + _Delivery.RecipAdr2
                                       + ":" + _Delivery.RecipAdr3
                                       + "+" + _Delivery.RecipCity
                                       + "++" + _Delivery.RecipZip
                                       + "+" + _Delivery.RecipCountryISOCode + "'";

                AppendObjectToEDIFile(_NAD3, typeof(NAD), filePath, "");


                CTA _CTA = new IFCSUMD96A.CTA();
                _CTA.Segment = "CTA"; // contact destinataire
                _CTA.Record = "++:" + _Delivery.RecipFirstName + "'";
                AppendObjectToEDIFile(_CTA, typeof(CTA), filePath, "");


                // Com -> if Phone number -> remove any << + >>

                string telephone = _Delivery.RecipPhoneNumber.Replace("+", "");

                COM _COM = new IFCSUMD96A.COM();
                _COM.Segment = "COM"; // numero telephone contact
                _COM.Record = "+" + telephone + ":TE" + "'";
                AppendObjectToEDIFile(_COM, typeof(COM), filePath, "");

                GID _GID = new IFCSUMD96A.GID();
                _GID.Segment = "GID"; //num Sequence + Numero UM + Type d'Unité de manutention
                _GID.Record = "+" + _Delivery.ParcelShipmentSequence + "+" + _DBSchenkerSettings.UMNumber + ":" + _Delivery.ShippingSupport + "'";
                AppendObjectToEDIFile(_GID, typeof(GID), filePath, "");

                MEA _MEA = new IFCSUMD96A.MEA();
                _MEA.Segment = "MEA";
                _MEA.Record = "+WT+AAE+KGM:" + _Delivery.ParcelWeight + "'";
                AppendObjectToEDIFile(_MEA, typeof(MEA), filePath, "");

                if (_Delivery.ParcelValue != null)
                {
                    MEA _MEA2 = new IFCSUMD96A.MEA();
                    _MEA2.Segment = "MEA";
                    _MEA2.Record = "+VOL+AAW+MTQ:" + _Delivery.ParcelValue + "'";
                    AppendObjectToEDIFile(_MEA2, typeof(MEA), filePath, "");
                }

                PCI _PCI = new IFCSUMD96A.PCI();
                _PCI.Segment = "PCI";
                _PCI.Record = "+18" + "'";
                AppendObjectToEDIFile(_PCI, typeof(PCI), filePath, "");


                // Get Barcode informations
                HeppnerEDI schenkerlabel = new HeppnerEDI();


                string Barcode = "*" + "219013" + _DBSchenkerSettings.ChargerCode + _Delivery.UniqueIdentifier + "3" + "00";

                GIN _GIN = new IFCSUMD96A.GIN();
                _GIN.Segment = "GIN"; // Valeur de Code Barres
                _GIN.Record = "+BN+" + Barcode + "'";
                AppendObjectToEDIFile(_GIN, typeof(GIN), filePath, "");
            }

            UNT _UNT = new IFCSUMD96A.UNT();
            // NOMBRE DE SEGMENTS A ENCORE DEFINIR
            _UNT.Segment = "UNT+33+" + _DBSchenkerSettings.MessageId + "'";
            AppendObjectToEDIFile(_UNT, typeof(UNT), filePath, "");

        }

        public void IFCSUMD96A_AddFooterToFile(List<DeliveryDM> _DeliveryDM, string filePath, DBSchenkerSettings _DBSchenkerSettings)
        {

            // Add footer

            UNZ _UNZ = new IFCSUMD96A.UNZ();
            _UNZ.Segment = "UNZ";
            _UNZ.Record = "+1+99999'";

            AppendObjectToEDIFile(_UNZ, typeof(UNZ), filePath, "");

        }

        public void AppendObjectToEDIFile(object _Record, Type RecordType, string path, string Delimiter)
        {

            var props = RecordType.GetProperties(BindingFlags.Public | BindingFlags.Instance);


            using (StreamWriter w = File.AppendText(path))
            {
                string _LineBuffer = string.Join(Delimiter, props.Select(p => p.GetValue(_Record, null)));

                // Get rid of unwanted characters

                switch (Delimiter)
                {
                    case "\t":
                        _LineBuffer = Regex.Replace(_LineBuffer, @"\n|\r", "");
                        break;
                    case "\n":
                        _LineBuffer = Regex.Replace(_LineBuffer, @"\t|\r", "");
                        break;
                    case "\r":
                        _LineBuffer = Regex.Replace(_LineBuffer, @"\t|\n", "");
                        break;
                    default:
                        _LineBuffer = Regex.Replace(_LineBuffer, @"\t|\n|\r", "");
                        break;
                }

                w.WriteLine(_LineBuffer);
            }
        }



        public class BGM
        {
            public string Segment { get; set; }
            public string Record { get; set; }

        }

        public class CNI
        {
            public string Segment { get; set; }
            public string Record { get; set; }

        }

        public class COM
        {
            public string Segment { get; set; }
            public string Record { get; set; }

        }

        public class CTA
        {
            public string Segment { get; set; }
            public string Record { get; set; }
        }

        public class FTX
        {
            public string Segment { get; set; }
            public string Record { get; set; }

        }

        public class DTM
        {
            public string Segment { get; set; }
            public string Record { get; set; }

        }

        public class GID
        {
            public string Segment { get; set; }
            public string Record { get; set; }
        }

        public class GIN
        {
            public string Segment { get; set; }
            public string Record { get; set; }

        }

        public class MEA
        {
            public string Segment { get; set; }
            public string Record { get; set; }
        }

        public class NAD
        {
            public string Segment { get; set; }
            public string Record { get; set; }

        }

        public class PCI
        {
            public string Segment { get; set; }
            public string Record { get; set; }

        }

        public class RFF
        {
            public string Segment { get; set; }
            public string Record { get; set; }

        }

        public class TDT
        {
            public string Segment { get; set; }
            public string Record { get; set; }

        }

        public class TOD
        {
            public string Segment { get; set; }
            public string Record { get; set; }
        }

        public class UNA
        {
            public string Segment { get; set; }
            public string Record { get; set; }

        }

        public class UNB
        {
            public string Segment { get; set; }
            public string Record { get; set; }
        }

        public class UNH
        {
            public string Segment { get; set; }
            public string Record { get; set; }

        }

        public class UNT
        {
            public string Segment { get; set; }

        }

        public class UNZ
        {
            public string Segment { get; set; }
            public string Record { get; set; }

        }



    }


}
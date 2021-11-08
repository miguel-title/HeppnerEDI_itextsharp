using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using Microsoft.VisualBasic.FileIO;

namespace AtomicSeller.ViewModels
{
    public class SchenkerPlanTransport
    {
        public class TransportVM
        {
            public string CodePays { get; set; }
            public string CodeVille { get; set; }
            public string Ville { get; set; }
            public string DateApplication { get; set; }
            public string CodeINSEE { get; set; }
            public string CodePlateforme1 { get; set; }
            public string CodePlateforme2 { get; set; }
            public string CodePlateforme3 { get; set; }
            public string CodePlateforme4 { get; set; }
            public string CodePlateforme5 { get; set; }
            public string AgenceLivraison { get; set; }
            public string ModeLivraison { get; set; }
            public string TourneeLivraison { get; set; }
            public string CodeDirectionnel { get; set; }
            public string EligibleWebService { get; set; }
            public string Reserve1 { get; set; }
            public string Reserve2 { get; set; }
            public string Reserve3 { get; set; }
            public string TypeDEtiquette { get; set; }
            public string TypeDeCodeABarres { get; set; }
            public string ContenuDuCodeABarres { get; set; }

        }

        public List<TransportVM> ReadTransportFromFile(string FileDir)
        {
            List<TransportVM> _TransportVMs = new List<TransportVM>();

            string FileName = "DBSC41";

            string FilePath = Path.Combine(FileDir, FileName);

            using (TextFieldParser parser = FileSystem.OpenTextFieldParser(FilePath))
            {
                // Set the field widths.
                parser.TextFieldType = FieldType.FixedWidth;
                parser.FieldWidths = new int[] { 3, 9, 35, 8, 9, 4, 4, 4, 4, 4, 4, 1, 4, 3, 1, 3, 4, 4, 4, 4, 4 };

                // Process the file's lines.
                while (!parser.EndOfData)
                {
                    try
                    {
                        int i = 0;
                        string[] _Fields = parser.ReadFields();

                        TransportVM _TransportVM = new TransportVM();

                        _TransportVM.CodePays = _Fields[i++];
                        _TransportVM.CodeVille = _Fields[i++];
                        _TransportVM.Ville = _Fields[i++];
                        _TransportVM.DateApplication = _Fields[i++];
                        _TransportVM.CodeINSEE = _Fields[i++];
                        _TransportVM.CodePlateforme1 = _Fields[i++];
                        _TransportVM.CodePlateforme2 = _Fields[i++];
                        _TransportVM.CodePlateforme3 = _Fields[i++];
                        _TransportVM.CodePlateforme4 = _Fields[i++];
                        _TransportVM.CodePlateforme5 = _Fields[i++];
                        _TransportVM.AgenceLivraison = _Fields[i++];
                        _TransportVM.ModeLivraison = _Fields[i++];
                        _TransportVM.TourneeLivraison = _Fields[i++];
                        _TransportVM.CodeDirectionnel = _Fields[i++];
                        _TransportVM.EligibleWebService = _Fields[i++];
                        _TransportVM.Reserve1 = _Fields[i++];
                        _TransportVM.Reserve2 = _Fields[i++];
                        _TransportVM.Reserve3 = _Fields[i++];
                        _TransportVM.TypeDEtiquette = _Fields[i++];
                        _TransportVM.TypeDeCodeABarres = _Fields[i++];
                        _TransportVM.ContenuDuCodeABarres = _Fields[i++];

                        _TransportVMs.Add(_TransportVM);

                    }
                    catch (Exception ex)
                    {
                    }
                }
            }
            return _TransportVMs;
        }

    }
}
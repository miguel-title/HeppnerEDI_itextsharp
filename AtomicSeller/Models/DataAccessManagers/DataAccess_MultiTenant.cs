using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.IO;
using System.Data.Entity;
using AtomicSeller.Models;
using AtomicSeller.ViewModels;
using AtomicSeller.Controllers;
using AtomicSeller.Helpers;
using Newtonsoft.Json;
using System.Web.Mvc;

namespace AtomicSeller
{
    public class DA_MT_TEST
    {
        public enum ShippingCarrierType : int
        {
            None = 0,
            Labels = 1,
            Enveloppes = 2,
            ExcelExport = 3,
            Colissimo = 4,
            MondialRelay = 5,
            UPS = 6,
            RoyalMail = 7,
            DHLExpress = 8,
            RelaisColis = 9,
            DHLParcel = 10
        }


        public static string CleanFieldForXML(string Field)
        {
            if (Field == null) return null;
            Field = Field.Replace(",", "");
            Field = Field.Replace("'", " ");
            Field = Field.Replace("&", "&amp;");
            Field = Field.Replace("<", "&lt;");
            Field = Field.Replace(">", "&gt;");
            Field = Field.Replace("\"", "&quot;");

            return (Field);

        }

        public static string CleanFieldForSQLServer(string Field)
        {
            //ajoute un anti - slash aux caractères suivants: NULL, \x00, \n, \r, \, ', " et \x1a.
            //if (Field.Contains ("ancienne"))
            //{
            //    MessageBox.Show(Field + "\n" + Field.Replace("'", "\\'"));
            //}
            if (Field == null) return null;
            Field = Field.Replace(",", "");
            Field = Field.Replace("'", " "); // !!!!!!!!!!
            Field = Field.Replace("  ", " ");
            Field = Field.Replace("\"", "");
            Field = Field.Replace("/", " ");
            return (Field);
            //return (Field.Replace("'", "\\\'"));
        }


        #region Order
        
        
        public static List<OrderDM> GetAllOrders()
        {
            return null;
        }
        
        public static List<OrderDM> GetOrdersByStatus(List<string> Order_StatusList, string StoreName)
        {
            return null;  
        }

   
        #endregion

        #region Shipment
        
        public static IEnumerable<string> ExtractValidationMessages(DbContext context)
        {
            var errors = context.GetValidationErrors();
            foreach (var error in errors)
            {
                yield return "Validation error on " + error.Entry.Entity.GetType().FullName + "{" + error.Entry.Entity.ToString() + "}";
                foreach (var errorMessage in error.ValidationErrors)
                {
                    yield return "Property " + errorMessage.PropertyName + " " + errorMessage.ErrorMessage;
                }
            }
        }

        public static string ValidationErrorMessage(DbContext context)
        {
            var toto = ExtractValidationMessages(context);
            string detail = string.Empty;

            foreach (var titi in toto) detail += titi + "\n";

            return detail;
        }




        

        public static List<DeliveryDM> GetAllShipments()
        {
            return null;
        }
        #endregion

     
        #region Orderproduct

        public static List<OrderProductDM> GetOrdersProductsByOrderID(int OrderID)
        {
            return null;
        }

        #endregion


        public static bool TestExistingShipment(int ShipmentID)
        {

            return false;
        }

        public static bool TestExistingOrderProduct(int OrderProductID)
        {

            return false;
        }

        public static int TestExistingOrder(string OrderKey, string OrderType=null, string Order_StoreType=null, string MerchantKey = null)
        {
            if (string.IsNullOrEmpty(OrderKey)) return 0;

            return 0;
        }





        

    }

}
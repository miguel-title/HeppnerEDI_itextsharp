using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using AtomicSeller.Models;

namespace ShopifyOrderAPI.Models
{
    public class GetOrdersRequest
    {
        public Header Header { get; set; }
        public GetOrdersData Request { get; set; }
    }

    public class PutOrderStatusRequest
    {
        public Header Header { get; set; }
        public PutOrderStatusData Request { get; set; }

    }
    public class Header
    {
        public string Token { get; set; }
    }

    public class GetOrdersData
    {
        public string OrderStartDate { get; set; }
        public string OrderEndDate { get; set; }
        public string LanguageCode { get; set; }
        public string OrderStatus { get; set; }

    }

    public class PutOrderStatusData
    {
        public string OrderID { get; set; }
        public string ShippingID { get; set; }
        public string LanguageCode { get; set; }
        public string TrackingNumber { get; set; }
        public string TrackingUrl { get; set; }

    }

    public class PutOrderStatusResponse
    {
        public ResponseHeader Header { get; set; }
    }

    public class GetOrdersResponse
    {
        public ResponseHeader Header { get; set; }
        public OrdersResponse Response { get; set; }
    }

    public class ResponseHeader
    {
        public string RequestStatus { get; set; }

        public string ReturnCode { get; set; }

        public string ReturnMessage { get; set; }

        public string LanguageCode { get; set; }
    }

    public class OrdersResponse
    {
        public List<Order> Orders { get; set; }
    }


    public class PutOrderParam
    {
        public string tracking_number { get; set; }
        public string id { get; set; }
    }

    public class Order
    {
        public string id { get; set; }
        public string email { get; set; }
        public object closed_at { get; set; }
        public DateTime created_at { get; set; }
        public DateTime updated_at { get; set; }
        public int number { get; set; }
        public object note { get; set; }
        public string token { get; set; }
        public string gateway { get; set; }
        public bool test { get; set; }
        public string total_price { get; set; }
        public string subtotal_price { get; set; }
        public int total_weight { get; set; }
        public string total_tax { get; set; }
        public bool taxes_included { get; set; }
        public string currency { get; set; }
        public string financial_status { get; set; }
        public bool confirmed { get; set; }
        public string total_discounts { get; set; }
        public string total_line_items_price { get; set; }
        public string cart_token { get; set; }
        public bool buyer_accepts_marketing { get; set; }
        public string name { get; set; }
        public string referring_site { get; set; }
        public string landing_site { get; set; }
        public object cancelled_at { get; set; }
        public object cancel_reason { get; set; }
        public string total_price_usd { get; set; }
        public string checkout_token { get; set; }
        public string reference { get; set; }
        public object user_id { get; set; }
        public object location_id { get; set; }
        public string source_identifier { get; set; }
        public object source_url { get; set; }
        public DateTime processed_at { get; set; }
        public object device_id { get; set; }
        public string phone { get; set; }
        public object customer_locale { get; set; }
        public object app_id { get; set; }
        public string browser_ip { get; set; }
        public string landing_site_ref { get; set; }
        public int order_number { get; set; }
        public Discount_Applications[] discount_applications { get; set; }
        public Discount_Codes[] discount_codes { get; set; }
        public Note_Attributes[] note_attributes { get; set; }
        public string[] payment_gateway_names { get; set; }
        public string processing_method { get; set; }
        public string checkout_id { get; set; }
        public string source_name { get; set; }
        public object fulfillment_status { get; set; }
        public Tax_Lines[] tax_lines { get; set; }
        public string tags { get; set; }
        public string contact_email { get; set; }
        public string order_status_url { get; set; }
        public string presentment_currency { get; set; }
        public Total_Line_Items_Price_Set total_line_items_price_set { get; set; }
        public Total_Discounts_Set total_discounts_set { get; set; }
        public Total_Shipping_Price_Set total_shipping_price_set { get; set; }
        public Subtotal_Price_Set subtotal_price_set { get; set; }
        public Total_Price_Set total_price_set { get; set; }
        public Total_Tax_Set total_tax_set { get; set; }
        public string admin_graphql_api_id { get; set; }
        public Line_Items[] line_items { get; set; }
        public Shipping_Lines[] shipping_lines { get; set; }
        public Billing_Address billing_address { get; set; }
        public Shipping_Address shipping_address { get; set; }
        public Fulfillment[] fulfillments { get; set; }
        public Client_Details client_details { get; set; }
        public Refund[] refunds { get; set; }
        public Payment_Details payment_details { get; set; }
        public Customer customer { get; set; }
    }

    public class Total_Line_Items_Price_Set
    {
        public Shop_Money shop_money { get; set; }
        public Presentment_Money presentment_money { get; set; }
    }

    public class Shop_Money
    {
        public string amount { get; set; }
        public string currency_code { get; set; }
    }

    public class Presentment_Money
    {
        public string amount { get; set; }
        public string currency_code { get; set; }
    }

    public class Total_Discounts_Set
    {
        public Shop_Money shop_money { get; set; }
        public Presentment_Money presentment_money { get; set; }
    }



    public class Total_Shipping_Price_Set
    {
        public Shop_Money shop_money { get; set; }
        public Presentment_Money presentment_money { get; set; }
    }


    public class Subtotal_Price_Set
    {
        public Shop_Money shop_money { get; set; }
        public Presentment_Money presentment_money { get; set; }
    }

    public class Total_Price_Set
    {
        public Shop_Money shop_money { get; set; }
        public Presentment_Money presentment_money { get; set; }
    }


    public class Total_Tax_Set
    {
        public Shop_Money shop_money { get; set; }
        public Presentment_Money presentment_money { get; set; }
    }


    public class Billing_Address
    {
        public string first_name { get; set; }
        public string address1 { get; set; }
        public string phone { get; set; }
        public string city { get; set; }
        public string zip { get; set; }
        public string province { get; set; }
        public string country { get; set; }
        public string last_name { get; set; }
        public string address2 { get; set; }
        public string company { get; set; }
        public float latitude { get; set; }
        public float longitude { get; set; }
        public string name { get; set; }
        public string country_code { get; set; }
        public string province_code { get; set; }
    }

    public class Shipping_Address
    {
        public string first_name { get; set; }
        public string address1 { get; set; }
        public string phone { get; set; }
        public string city { get; set; }
        public string zip { get; set; }
        public string province { get; set; }
        public string country { get; set; }
        public string last_name { get; set; }
        public string address2 { get; set; }
        public object company { get; set; }
        public string latitude { get; set; }
        public string longitude { get; set; }
        public string name { get; set; }
        public string country_code { get; set; }
        public string province_code { get; set; }
    }

    public class Client_Details
    {
        public string browser_ip { get; set; }
        public object accept_language { get; set; }
        public object user_agent { get; set; }
        public object session_hash { get; set; }
        public object browser_width { get; set; }
        public object browser_height { get; set; }
    }

    public class Payment_Details
    {
        public object credit_card_bin { get; set; }
        public object avs_result_code { get; set; }
        public object cvv_result_code { get; set; }
        public string credit_card_number { get; set; }
        public string credit_card_company { get; set; }
    }

    public class Customer
    {
        public string id { get; set; }
        public string email { get; set; }
        public bool accepts_marketing { get; set; }
        public DateTime created_at { get; set; }
        public DateTime updated_at { get; set; }
        public string first_name { get; set; }
        public string last_name { get; set; }
        public int orders_count { get; set; }
        public string state { get; set; }
        public string total_spent { get; set; }
        public int last_order_id { get; set; }
        public object note { get; set; }
        public bool verified_email { get; set; }
        public object multipass_identifier { get; set; }
        public bool tax_exempt { get; set; }
        public object phone { get; set; }
        public string tags { get; set; }
        public string last_order_name { get; set; }
        public string currency { get; set; }
        public DateTime accepts_marketing_updated_at { get; set; }
        public object marketing_opt_in_level { get; set; }
        public object[] tax_exemptions { get; set; }
        public string admin_graphql_api_id { get; set; }
        public Default_Address default_address { get; set; }
    }

    public class Default_Address
    {
        public string id { get; set; }
        public string customer_id { get; set; }
        public object first_name { get; set; }
        public object last_name { get; set; }
        public object company { get; set; }
        public string address1 { get; set; }
        public string address2 { get; set; }
        public string city { get; set; }
        public string province { get; set; }
        public string country { get; set; }
        public string zip { get; set; }
        public string phone { get; set; }
        public string name { get; set; }
        public string province_code { get; set; }
        public string country_code { get; set; }
        public string country_name { get; set; }
        public bool _default { get; set; }
    }

    public class Discount_Applications
    {
        public string type { get; set; }
        public string value { get; set; }
        public string value_type { get; set; }
        public string allocation_method { get; set; }
        public string target_selection { get; set; }
        public string target_type { get; set; }
        public string code { get; set; }
    }

    public class Discount_Codes
    {
        public string code { get; set; }
        public string amount { get; set; }
        public string type { get; set; }
    }

    public class Note_Attributes
    {
        public string name { get; set; }
        public string value { get; set; }
    }

    public class Tax_Lines
    {
        public string price { get; set; }
        public float rate { get; set; }
        public string title { get; set; }
        public Price_Set price_set { get; set; }
    }

    public class Price_Set
    {
        public Shop_Money shop_money { get; set; }
        public Presentment_Money presentment_money { get; set; }
    }


    public class Line_Items
    {
        public string id { get; set; }
        public string variant_id { get; set; }
        public string title { get; set; }
        public int quantity { get; set; }
        public string sku { get; set; }
        public string variant_title { get; set; }
        public object vendor { get; set; }
        public string fulfillment_service { get; set; }
        public string product_id { get; set; }
        public bool requires_shipping { get; set; }
        public bool taxable { get; set; }
        public bool gift_card { get; set; }
        public string name { get; set; }
        public string variant_inventory_management { get; set; }
        public Property[] properties { get; set; }
        public bool product_exists { get; set; }
        public int fulfillable_quantity { get; set; }
        public int grams { get; set; }
        public string price { get; set; }
        public string total_discount { get; set; }
        public object fulfillment_status { get; set; }
        public Price_Set price_set { get; set; }
        public Total_Discount_Set total_discount_set { get; set; }
        public object[] discount_allocations { get; set; }
        public string admin_graphql_api_id { get; set; }
        public Tax_Lines[] tax_lines { get; set; }
    }


    public class Total_Discount_Set
    {
        public Shop_Money shop_money { get; set; }
        public Presentment_Money presentment_money { get; set; }
    }

    public class Shipping_Lines
    {
        public string id { get; set; }
        public string title { get; set; }
        public string price { get; set; }
        public string code { get; set; }
        public string source { get; set; }
        public object phone { get; set; }
        public object requested_fulfillment_service_id { get; set; }
        public object delivery_category { get; set; }
        public object carrier_identifier { get; set; }
        public string discounted_price { get; set; }
        public Price_Set price_set { get; set; }
        public Discounted_Price_Set discounted_price_set { get; set; }
        public object[] discount_allocations { get; set; }
        public object[] tax_lines { get; set; }
    }

    public class Discounted_Price_Set
    {
        public Shop_Money shop_money { get; set; }
        public Presentment_Money presentment_money { get; set; }
    }


    public class Fulfillment
    {
        public string id { get; set; }
        public string order_id { get; set; }
        public string status { get; set; }
        public DateTime created_at { get; set; }
        public string service { get; set; }
        public DateTime updated_at { get; set; }
        public object tracking_company { get; set; }
        public object shipment_status { get; set; }
        public string location_id { get; set; }
        public string tracking_number { get; set; }
        public string[] tracking_numbers { get; set; }
        public string tracking_url { get; set; }
        public string[] tracking_urls { get; set; }
        public Receipt receipt { get; set; }
        public string name { get; set; }
        public string admin_graphql_api_id { get; set; }
        public Line_Items[] line_items { get; set; }
    }

    public class Receipt
    {
        public bool testcase { get; set; }
        public string authorization { get; set; }
    }




    public class Refund
    {
        public string id { get; set; }
        public string order_id { get; set; }
        public DateTime created_at { get; set; }
        public string note { get; set; }
        public string user_id { get; set; }
        public DateTime processed_at { get; set; }
        public bool restock { get; set; }
        public string admin_graphql_api_id { get; set; }
        public Refund_Line_Items[] refund_line_items { get; set; }
        public Transaction[] transactions { get; set; }
        public object[] order_adjustments { get; set; }
    }

    public class Refund_Line_Items
    {
        public string id { get; set; }
        public int quantity { get; set; }
        public int line_item_id { get; set; }
        public int location_id { get; set; }
        public string restock_type { get; set; }
        public float subtotal { get; set; }
        public float total_tax { get; set; }
        public Subtotal_Set subtotal_set { get; set; }
        public Total_Tax_Set total_tax_set { get; set; }
        public Line_Item line_item { get; set; }
    }

    public class Subtotal_Set
    {
        public Shop_Money shop_money { get; set; }
        public Presentment_Money presentment_money { get; set; }
    }


    public class Line_Item
    {
        public string id { get; set; }
        public string variant_id { get; set; }
        public string title { get; set; }
        public int quantity { get; set; }
        public string sku { get; set; }
        public string variant_title { get; set; }
        public object vendor { get; set; }
        public string fulfillment_service { get; set; }
        public string product_id { get; set; }
        public bool requires_shipping { get; set; }
        public bool taxable { get; set; }
        public bool gift_card { get; set; }
        public string name { get; set; }
        public string variant_inventory_management { get; set; }
        public Property[] properties { get; set; }
        public bool product_exists { get; set; }
        public int fulfillable_quantity { get; set; }
        public int grams { get; set; }
        public string price { get; set; }
        public string total_discount { get; set; }
        public object fulfillment_status { get; set; }
        public Price_Set price_set { get; set; }
        public Total_Discount_Set total_discount_set { get; set; }
        public object[] discount_allocations { get; set; }
        public string admin_graphql_api_id { get; set; }
        public Tax_Lines[] tax_lines { get; set; }
    }



    public class Property
    {
        public string name { get; set; }
        public string value { get; set; }
    }


    public class Transaction
    {
        public string id { get; set; }
        public string order_id { get; set; }
        public string kind { get; set; }
        public string gateway { get; set; }
        public string status { get; set; }
        public object message { get; set; }
        public DateTime created_at { get; set; }
        public bool test { get; set; }
        public string authorization { get; set; }
        public object location_id { get; set; }
        public object user_id { get; set; }
        public int parent_id { get; set; }
        public DateTime processed_at { get; set; }
        public object device_id { get; set; }
        public Receipt receipt { get; set; }
        public object error_code { get; set; }
        public string source_name { get; set; }
        public string amount { get; set; }
        public string currency { get; set; }
        public string admin_graphql_api_id { get; set; }
    }



}
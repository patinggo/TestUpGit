using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Services;

using SAP.Middleware.Connector;
using System.Data;
using System.Collections;

//[WebService(Namespace = "http://192.168.5.109/GZWebService/")]
[WebService(Namespace = "http://tempuri.org/")]
[WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
// To allow this Web Service to be called from script, using ASP.NET AJAX, uncomment the following line. 
// [System.Web.Script.Services.ScriptService]

public class GZSAPService : System.Web.Services.WebService
{
    public GZSAPService()
    {

        //Uncomment the following line if using designed components 
        //InitializeComponent(); 
    }

    DataTable dtReturn = new DataTable("dtReturn");
    DataTable dtStock = new DataTable("dtStock");
    DataTable dtException = new DataTable("dtException");

    DataTable dtSOReturn = new DataTable("dtSOReturn");
    DataTable dtSODoc = new DataTable("dtSODoc");
    DataTable dtSOException = new DataTable("dtSOException");

    DataTable dtCreditLimit = new DataTable("dtCreditLimit");


    //Check Stock
    #region "Check Stock"
    [WebMethod]
    public DataSet GetStock(string materialId)
    {
        RfcConfigParameters rfc = new RfcConfigParameters();
        rfc.Add(RfcConfigParameters.Name, "CheckStock");
        rfc.Add(RfcConfigParameters.AppServerHost, "192.168.2.51");
        rfc.Add(RfcConfigParameters.Client, "400");
        rfc.Add(RfcConfigParameters.User, "basis");
        rfc.Add(RfcConfigParameters.Password, "support32");
        rfc.Add(RfcConfigParameters.SystemNumber, "00");
        rfc.Add(RfcConfigParameters.Language, "EN");
        //rfc.Add(RfcConfigParameters.PoolSize, "5");
        //Server Develop
        //rfc.Add(RfcConfigParameters.AppServerHost, "192.168.2.53");
        //rfc.Add(RfcConfigParameters.Client, "220");
        //Server Production
        //rfc.Add(RfcConfigParameters.AppServerHost, "192.168.2.52");
        //rfc.Add(RfcConfigParameters.Client, "300");
        //Server Real
        //rfc.Add(RfcConfigParameters.AppServerHost, "192.168.2.51");
        //rfc.Add(RfcConfigParameters.Client, "400");
        //rfc.Add(RfcConfigParameters.User, "basis");
        //rfc.Add(RfcConfigParameters.Password, "support32");
        //rfc.Add(RfcConfigParameters.SystemNumber, "00");
        //rfc.Add(RfcConfigParameters.Language, "TH");

        RfcDestination rfcDest = RfcDestinationManager.GetDestination(rfc);
        RfcRepository rfcRep = rfcDest.Repository;
        IRfcFunction function = rfcRep.CreateFunction("ZIMM0002");

        function["I_MATERIAL"].SetValue(materialId);
        function["I_PLANT"].SetValue("");
        function["I_SPLIT_PLANT"].SetValue("");

        Stock_CreateDatatable();
        try
        {
            //execute the function
            function.Invoke(rfcDest);
            dtException.Rows.Add("Y");
        }
        catch (Exception e)
        {
            dtException.Rows.Add(e.Message);
        }

        IRfcTable tbReturn = function.GetTable("ET_RETURN");
        IRfcTable tbStock = function.GetTable("ET_STOCK");

        //Create Dataset
        return Stock_CreateDataset(tbReturn, tbStock, dtException);
    }
    [WebMethod]
    public string GetStock2(string materialId)
    {
        
        RfcConfigParameters rfc = new RfcConfigParameters();
        rfc.Add(RfcConfigParameters.Name, "CheckStock");
        rfc.Add(RfcConfigParameters.AppServerHost, "192.168.2.51");
        rfc.Add(RfcConfigParameters.Client, "400");
        rfc.Add(RfcConfigParameters.User, "basis");
        rfc.Add(RfcConfigParameters.Password, "support32");
        rfc.Add(RfcConfigParameters.SystemNumber, "00");
        rfc.Add(RfcConfigParameters.Language, "EN");
        //rfc.Add(RfcConfigParameters.AppServerHost, "192.168.2.53");
        //rfc.Add(RfcConfigParameters.Client, "220");
        //rfc.Add(RfcConfigParameters.User, "basis");
        //rfc.Add(RfcConfigParameters.Password, "demo2013");
        //rfc.Add(RfcConfigParameters.SystemNumber, "00");
        //rfc.Add(RfcConfigParameters.Language, "EN");
        //rfc.Add(RfcConfigParameters.PoolSize, "5");

        //Production
        //rfc.Add(RfcConfigParameters.Name, "CheckStock");
        //rfc.Add(RfcConfigParameters.AppServerHost, "192.168.2.51");
        //rfc.Add(RfcConfigParameters.Client, "400");
        //rfc.Add(RfcConfigParameters.User, "basis");
        //rfc.Add(RfcConfigParameters.Password, "support32");
        //rfc.Add(RfcConfigParameters.SystemNumber, "00");
        //rfc.Add(RfcConfigParameters.Language, "EN");

        RfcDestination rfcDest = RfcDestinationManager.GetDestination(rfc);
        RfcRepository rfcRep = rfcDest.Repository;
        IRfcFunction function = rfcRep.CreateFunction("ZIMM0002");

        function["I_MATERIAL"].SetValue(materialId);
        function["I_PLANT"].SetValue("");
        function["I_SPLIT_PLANT"].SetValue("");

        try
        {
            //execute the function
            function.Invoke(rfcDest);
        }
        catch (Exception e)
        {
            //Console.WriteLine(e.Message);
            return e.Message;
        }

        IRfcTable tbReturn = function.GetTable("ET_RETURN");
        IRfcTable tb = function.GetTable("ET_STOCK");

        string rType = tbReturn.GetString("TYPE");

        string sku = tb.GetString("MATERIAL");
        string Plant = tb.GetString("PLANT");
        string UOM = tb.GetString("BASE_Uom");
        double URStock = Convert.ToDouble(tb.GetString("UNRESTRICTED_STCK"));
        double OpenPO = Convert.ToDouble(tb.GetString("OPEN_PO"));
        double OpenSO = Convert.ToDouble(tb.GetString("OPEN_SO"));
        double OpenDO = Convert.ToDouble(tb.GetString("OPEN_DO"));

        string result = URStock.ToString(); return result;

    }
    public void Stock_CreateDatatable()
    {
        //dtException
        dtException.Columns.Add("Message");
        dtException.Clear();

        //dtReturn
        dtReturn.Columns.Add("Type"); dtReturn.Columns.Add("Id"); dtReturn.Columns.Add("Number");
        dtReturn.Columns.Add("Message"); dtReturn.Columns.Add("Log_No"); dtReturn.Columns.Add("Log_Msg_No");
        dtReturn.Columns.Add("Message_V1"); dtReturn.Columns.Add("Message_V2"); 
        dtReturn.Columns.Add("Message_V3"); dtReturn.Columns.Add("Message_V4");
        dtReturn.Columns.Add("Parameter"); dtReturn.Columns.Add("Row"); dtReturn.Columns.Add("Field"); dtReturn.Columns.Add("SYSTEM");
        dtReturn.Clear();

        //dtStock
        dtStock.Columns.Add("MATERIAL"); dtStock.Columns.Add("PLANT"); dtStock.Columns.Add("BASE_Uom");
        dtStock.Columns.Add("UNRESTRICTED_STCK"); dtStock.Columns.Add("OPEN_PO");
        dtStock.Columns.Add("OPEN_SO"); dtStock.Columns.Add("OPEN_DO");
        dtStock.Clear();

    }
    public DataSet Stock_CreateDataset(IRfcTable tbReturn, IRfcTable tbStock, DataTable dtException)
    {
        //tbReturn
        for(int i =0; i < tbReturn.RowCount; i++)
        {
            tbReturn.CurrentIndex = i;
            dtReturn.Rows.Add
                (
                tbReturn.GetString("Type"),tbReturn.GetString("Id"),tbReturn.GetString("Number"),
                tbReturn.GetString("Message"), tbReturn.GetString("Log_No"), tbReturn.GetString("Log_Msg_No"),
                tbReturn.GetString("Message_V1"), tbReturn.GetString("Message_V2"), 
                tbReturn.GetString("Message_V3"), tbReturn.GetString("Message_V4"),
                tbReturn.GetString("Parameter"), tbReturn.GetString("Row"), tbReturn.GetString("Field"), tbReturn.GetString("SYSTEM")
                );
        }

        //tbStock
        for (int j = 0; j < tbStock.RowCount; j++)
        {
            tbStock.CurrentIndex = j;
            dtStock.Rows.Add
                (
                tbStock.GetString("MATERIAL"), tbStock.GetString("PLANT"), tbStock.GetString("BASE_Uom"),
                tbStock.GetString("UNRESTRICTED_STCK"), tbStock.GetString("OPEN_PO"),
                tbStock.GetString("OPEN_SO"), tbStock.GetString("OPEN_DO")
                );
        }

        DataSet dsResult = new DataSet("dsResult");
        dsResult.Tables.Add(dtException);
        dsResult.Tables.Add(dtReturn);
        dsResult.Tables.Add(dtStock);

        return dsResult;
    }
    #endregion


    //Check Stock by Store
    #region "Check Stock by Store"
    [WebMethod]
    public DataSet GetStockByStore(string materialId,string storeId)
    {
        RfcConfigParameters rfc = new RfcConfigParameters();
        rfc.Add(RfcConfigParameters.Name, "CheckStock");
        rfc.Add(RfcConfigParameters.AppServerHost, "192.168.2.51");
        rfc.Add(RfcConfigParameters.Client, "400");
        rfc.Add(RfcConfigParameters.User, "basis");
        rfc.Add(RfcConfigParameters.Password, "support32");
        rfc.Add(RfcConfigParameters.SystemNumber, "00");
        rfc.Add(RfcConfigParameters.Language, "EN");
        //rfc.Add(RfcConfigParameters.PoolSize, "5");
        //Server Develop
        //rfc.Add(RfcConfigParameters.AppServerHost, "192.168.2.53");
        //rfc.Add(RfcConfigParameters.Client, "220");
        //Server Production
        //rfc.Add(RfcConfigParameters.AppServerHost, "192.168.2.52");
        //rfc.Add(RfcConfigParameters.Client, "300");
        //Server Real
        //rfc.Add(RfcConfigParameters.AppServerHost, "192.168.2.51");
        //rfc.Add(RfcConfigParameters.Client, "400");
        //rfc.Add(RfcConfigParameters.User, "basis");
        //rfc.Add(RfcConfigParameters.Password, "support32");
        //rfc.Add(RfcConfigParameters.SystemNumber, "00");
        //rfc.Add(RfcConfigParameters.Language, "TH");

        RfcDestination rfcDest = RfcDestinationManager.GetDestination(rfc);
        RfcRepository rfcRep = rfcDest.Repository;
        IRfcFunction function = rfcRep.CreateFunction("ZIMM0002");

        function["I_MATERIAL"].SetValue(materialId);
        function["I_PLANT"].SetValue(storeId);
        function["I_SPLIT_PLANT"].SetValue("");

        Stock_CreateDatatable();
        try
        {
            //execute the function
            function.Invoke(rfcDest);
            dtException.Rows.Add("Y");
        }
        catch (Exception e)
        {
            dtException.Rows.Add(e.Message);
        }

        IRfcTable tbReturn = function.GetTable("ET_RETURN");
        IRfcTable tbStock = function.GetTable("ET_STOCK");

        //Create Dataset
        return Stock_CreateDataset(tbReturn, tbStock, dtException);
    }
    [WebMethod]
    public string GetStockByStore2(string materialId, string storeId)
    {

        RfcConfigParameters rfc = new RfcConfigParameters();
        rfc.Add(RfcConfigParameters.Name, "CheckStock");
        rfc.Add(RfcConfigParameters.AppServerHost, "192.168.2.51");
        rfc.Add(RfcConfigParameters.Client, "400");
        rfc.Add(RfcConfigParameters.User, "basis");
        rfc.Add(RfcConfigParameters.Password, "support32");
        rfc.Add(RfcConfigParameters.SystemNumber, "00");
        rfc.Add(RfcConfigParameters.Language, "EN");

        RfcDestination rfcDest = RfcDestinationManager.GetDestination(rfc);
        RfcRepository rfcRep = rfcDest.Repository;
        IRfcFunction function = rfcRep.CreateFunction("ZIMM0002");

        function["I_MATERIAL"].SetValue(materialId);
        function["I_PLANT"].SetValue(storeId);
        function["I_SPLIT_PLANT"].SetValue("");

        try
        {
            //execute the function
            function.Invoke(rfcDest);
        }
        catch (Exception e)
        {
            //Console.WriteLine(e.Message);
            return e.Message;
        }

        IRfcTable tbReturn = function.GetTable("ET_RETURN");
        IRfcTable tb = function.GetTable("ET_STOCK");
        string result = "";
        if (tbReturn.RowCount == 0)
        {
            result = "materialId: " + materialId + " is not maintain in store[storeId: " + storeId + "].";
        }
        else
        { 
            string rType = tbReturn.GetString("TYPE");

            string sku = tb.GetString("MATERIAL");
            string Plant = tb.GetString("PLANT");
            string UOM = tb.GetString("BASE_Uom");
            double URStock = Convert.ToDouble(tb.GetString("UNRESTRICTED_STCK"));
            double OpenPO = Convert.ToDouble(tb.GetString("OPEN_PO"));
            double OpenSO = Convert.ToDouble(tb.GetString("OPEN_SO"));
            double OpenDO = Convert.ToDouble(tb.GetString("OPEN_DO"));

           result = URStock.ToString(); 
        }
        return result;


    }
    #endregion


    #region  check Credit limit


    [WebMethod]
    public DataSet CheckCreditlimit(string CreditControl, string customerNo)
    {
        RfcConfigParameters rfc = new RfcConfigParameters();

        rfc.Add(RfcConfigParameters.AppServerHost, "192.168.2.51");
        rfc.Add(RfcConfigParameters.Client, "400");
        rfc.Add(RfcConfigParameters.Name, "ZISD0002");
        rfc.Add(RfcConfigParameters.User, "basis");
        rfc.Add(RfcConfigParameters.Password, "support32");
        rfc.Add(RfcConfigParameters.SystemNumber, "00");
        rfc.Add(RfcConfigParameters.Language, "EN");


        RfcDestination rfcDest = RfcDestinationManager.GetDestination(rfc);
        RfcRepository rfcRep = rfcDest.Repository;
        IRfcFunction function = rfcRep.CreateFunction("ZISD0002");

        function["I_CRE_CONTRL_AREA"].SetValue(CreditControl);
        function["I_CUS_NO"].SetValue(customerNo);

        CreditLimit_CreateDatatable();
        try
        {
            //execute the function
            function.Invoke(rfcDest);
            dtException.Rows.Add("Y");
        }
        catch (Exception e)
        {
            dtException.Rows.Add(e.Message);
        }

        //DataRow dtCreditLimit = new DataRow();
        dtCreditLimit.Rows.Add(
                function.GetValue("E_CREDIT_ACCOUNT"),
                function.GetValue("E_CREDIT_EXPOSURE"),
                function.GetValue("E_CREDIT_LIMIT"),
                function.GetValue("E_LIABILITIES"),
                function.GetValue("E_PERCENTAGE"),
                function.GetValue("E_RECEIVABLES"),
                function.GetValue("E_SALES_VALUE"),
                function.GetValue("I_CRE_CONTRL_AREA"),
                function.GetValue("I_CUS_NO")
            );

        IRfcTable tbReturnu = function.GetTable("T_RETURN");


        if (tbReturnu.Count() > 0)
        {
            dtReturn.Rows.Add
               (
               tbReturnu.GetString("TYPE"), tbReturnu.GetString("ID"),
               tbReturnu.GetString("NUMBER"), tbReturnu.GetString("MESSAGE"),
               tbReturnu.GetString("LOG_NO"), tbReturnu.GetString("LOG_MSG_NO"),
               tbReturnu.GetString("MESSAGE_V1"), tbReturnu.GetString("MESSAGE_V2"),
               tbReturnu.GetString("MESSAGE_V3"), tbReturnu.GetString("MESSAGE_V4"),
               tbReturnu.GetString("PARAMETER"), tbReturnu.GetString("ROW"),
               tbReturnu.GetString("FIELD"), tbReturnu.GetString("SYSTEM")
               );
        }






        //create Dataset
        DataSet dsResultCredit = new DataSet("dsResultCredit");
        dsResultCredit.Tables.Add(dtException);
        dsResultCredit.Tables.Add(dtCreditLimit);
        dsResultCredit.Tables.Add(dtReturn);
        //Create Dataset
        return dsResultCredit;
    }



    public void CreditLimit_CreateDatatable()
    {
        //dtException
        dtException.Columns.Add("Message");
        dtException.Clear();

        //CreditLimit
        dtCreditLimit.Columns.Add("E_CREDIT_ACCOUNT"); dtCreditLimit.Columns.Add("E_CREDIT_EXPOSURE");
        dtCreditLimit.Columns.Add("E_CREDIT_LIMIT"); dtCreditLimit.Columns.Add("E_LIABILITIES");
        dtCreditLimit.Columns.Add("E_PERCENTAGE"); dtCreditLimit.Columns.Add("E_RECEIVABLES");
        dtCreditLimit.Columns.Add("E_SALES_VALUE"); dtCreditLimit.Columns.Add("I_CRE_CONTRL_AREA");
        dtCreditLimit.Columns.Add("I_CUS_NO");
        dtCreditLimit.Clear();


        //dtReturn   
        dtReturn.Columns.Add("TYPE"); dtReturn.Columns.Add("ID");
        dtReturn.Columns.Add("NUMBER"); dtReturn.Columns.Add("MESSAGE");
        dtReturn.Columns.Add("LOG_NO"); dtReturn.Columns.Add("LOG_MSG_NO");
        dtReturn.Columns.Add("MESSAGE_V1"); dtReturn.Columns.Add("MESSAGE_V2");
        dtReturn.Columns.Add("MESSAGE_V3"); dtReturn.Columns.Add("MESSAGE_V4");
        dtReturn.Columns.Add("PARAMETER"); dtReturn.Columns.Add("ROW");
        dtReturn.Columns.Add("FIELD"); dtReturn.Columns.Add("SYSTEM");
        dtReturn.Clear();

    }
    #endregion



    //Create SO Doc
    #region "Create SO Doc"
    [WebMethod]
    public DataSet CreateSODoc(DataSet dsData)
    {
        RfcConfigParameters rfc = new RfcConfigParameters();

        rfc.Add(RfcConfigParameters.AppServerHost, "192.168.2.51");
        rfc.Add(RfcConfigParameters.Client, "400");
        rfc.Add(RfcConfigParameters.Name, "CreateSODoc");
        rfc.Add(RfcConfigParameters.User, "basis");
        rfc.Add(RfcConfigParameters.Password, "support32");
        rfc.Add(RfcConfigParameters.SystemNumber, "00");
        rfc.Add(RfcConfigParameters.Language, "EN");


        RfcDestination rfcDest = RfcDestinationManager.GetDestination(rfc);
        RfcRepository rfcRep = rfcDest.Repository;
        IRfcFunction function = rfcRep.CreateFunction("ZISD0001_POST_SO_DOC");

        function["I_TEST"].SetValue("");

        //1)Zisd0001HdSoShipTo
        DataTable dtShipTo = dsData.Tables["dtShipTo"];
        IRfcStructure tbShipTo = function.GetStructure("IS_ADDR_SHIP_TO");
        //IRfcTable tbShipTo = function.GetTable("IS_ADDR_SHIP_TO");
        //tbShipTo.Append();      
        tbShipTo.SetValue("NAME", dtShipTo.Rows[0]["NAME"]);
        tbShipTo.SetValue("NAME_2", dtShipTo.Rows[0]["NAME_2"]);
        tbShipTo.SetValue("NAME_3", dtShipTo.Rows[0]["NAME_3"]);
        tbShipTo.SetValue("NAME_4", dtShipTo.Rows[0]["NAME_4"]);
        tbShipTo.SetValue("STREET_LNG", dtShipTo.Rows[0]["STREET_LNG"]);
        tbShipTo.SetValue("STR_SUPPL1", dtShipTo.Rows[0]["STR_SUPPL1"]);
        tbShipTo.SetValue("DISTRICT", dtShipTo.Rows[0]["DISTRICT"]);
        tbShipTo.SetValue("CITY", dtShipTo.Rows[0]["CITY"]);
        tbShipTo.SetValue("POSTL_COD1", dtShipTo.Rows[0]["POSTL_COD1"]);
        tbShipTo.SetValue("COUNTRY", dtShipTo.Rows[0]["COUNTRY"]);
        tbShipTo.SetValue("TEL1_NUMBR", dtShipTo.Rows[0]["TEL1_NUMBR"]);
        tbShipTo.SetValue("TEL1_EXT", dtShipTo.Rows[0]["TEL1_EXT"]);
        tbShipTo.SetValue("FAX_NUMBER", dtShipTo.Rows[0]["FAX_NUMBER"]);
        tbShipTo.SetValue("FAX_EXTENS", dtShipTo.Rows[0]["FAX_EXTENS"]);
        tbShipTo.SetValue("E_MAIL", dtShipTo.Rows[0]["E_MAIL"]);
        tbShipTo.SetValue("CONTACT_PERSON", dtShipTo.Rows[0]["CONTACT_PERSON"]);
        tbShipTo.SetValue("TRANSPZONE", dtShipTo.Rows[0]["TRANSPZONE"]);
        function.SetValue("IS_ADDR_SHIP_TO", tbShipTo);

        //2)Zisd0001HdSoSoldTo
        DataTable dtSoldTo = dsData.Tables["dtSoldTo"];
        IRfcStructure tbSoldTo = function.GetStructure("IS_ADDR_SOLD_TO");
        //IRfcTable tbSoldTo = function.GetTable("IS_ADDR_SOLD_TO");
        //tbSoldTo.Append();
        tbSoldTo.SetValue("NAME", dtSoldTo.Rows[0]["NAME"]);
        tbSoldTo.SetValue("NAME_2", dtSoldTo.Rows[0]["NAME_2"]);
        tbSoldTo.SetValue("NAME_3", dtSoldTo.Rows[0]["NAME_3"]);
        tbSoldTo.SetValue("NAME_4", dtSoldTo.Rows[0]["NAME_4"]);
        tbSoldTo.SetValue("STREET_LNG", dtSoldTo.Rows[0]["STREET_LNG"]);
        tbSoldTo.SetValue("STR_SUPPL1", dtSoldTo.Rows[0]["STR_SUPPL1"]);
        tbSoldTo.SetValue("DISTRICT", dtSoldTo.Rows[0]["DISTRICT"]);
        tbSoldTo.SetValue("CITY", dtSoldTo.Rows[0]["CITY"]);
        tbSoldTo.SetValue("POSTL_COD1", dtSoldTo.Rows[0]["POSTL_COD1"]);
        tbSoldTo.SetValue("COUNTRY", dtSoldTo.Rows[0]["COUNTRY"]);
        tbSoldTo.SetValue("TEL1_NUMBR", dtSoldTo.Rows[0]["TEL1_NUMBR"]);
        tbSoldTo.SetValue("TEL1_EXT", dtSoldTo.Rows[0]["TEL1_EXT"]);
        tbSoldTo.SetValue("FAX_NUMBER", dtSoldTo.Rows[0]["FAX_NUMBER"]);
        tbSoldTo.SetValue("FAX_EXTENS", dtSoldTo.Rows[0]["FAX_EXTENS"]);
        tbSoldTo.SetValue("E_MAIL", dtSoldTo.Rows[0]["E_MAIL"]);
        tbSoldTo.SetValue("CONTACT_PERSON", dtSoldTo.Rows[0]["CONTACT_PERSON"]);
        tbSoldTo.SetValue("TAX_ID", dtSoldTo.Rows[0]["TAX_ID"]);
        tbSoldTo.SetValue("BRANCH_NO", dtSoldTo.Rows[0]["BRANCH_NO"]);
        function.SetValue("IS_ADDR_SOLD_TO", tbSoldTo);

        //3)Zisd0001HdFi(ส่วนมัดจำ)
        DataTable dtHdFi = dsData.Tables["dtHdFi"];
        IRfcStructure tbHdFi = function.GetStructure("IS_HD_FI");
        //IRfcTable tbHdFi = function.GetTable("IS_HD_FI");
        //tbHdFi.Append();
        tbHdFi.SetValue("BLDAT", dtHdFi.Rows[0]["BLDAT"]);
        tbHdFi.SetValue("NEWKO", dtHdFi.Rows[0]["NEWKO"]);
        tbHdFi.SetValue("FWBAS", dtHdFi.Rows[0]["FWBAS"]);
        tbHdFi.SetValue("WRBTR", dtHdFi.Rows[0]["WRBTR"]);
        function.SetValue("IS_HD_FI", tbHdFi);

        //4)Zisd0001HdSo
        DataTable dtHdSo = dsData.Tables["dtHdSo"];
        IRfcStructure tbHdSo = function.GetStructure("IS_HD_SO");
        //IRfcTable tbHdSo = function.GetTable("IS_HD_SO");
        //tbHdSo.Append();
        tbHdSo.SetValue("HD_INDICATOR", dtHdSo.Rows[0]["HD_INDICATOR"]);
        tbHdSo.SetValue("PURCH_NO_S", dtHdSo.Rows[0]["PURCH_NO_S"]);
        tbHdSo.SetValue("PO_DAT_S", dtHdSo.Rows[0]["PO_DAT_S"]);
        tbHdSo.SetValue("DOC_TYPE", dtHdSo.Rows[0]["DOC_TYPE"]);
        tbHdSo.SetValue("SALES_ORG", dtHdSo.Rows[0]["SALES_ORG"]);
        tbHdSo.SetValue("DISTR_CHAN", dtHdSo.Rows[0]["DISTR_CHAN"]);
        tbHdSo.SetValue("DIVISION", dtHdSo.Rows[0]["DIVISION"]);
        tbHdSo.SetValue("REQ_DATE_H", dtHdSo.Rows[0]["REQ_DATE_H"]);
        tbHdSo.SetValue("PURCH_NO_C", dtHdSo.Rows[0]["PURCH_NO_C"]);
        tbHdSo.SetValue("PLANT", dtHdSo.Rows[0]["PLANT"]);
        tbHdSo.SetValue("SHIP_COND", dtHdSo.Rows[0]["SHIP_COND"]);
        tbHdSo.SetValue("CURRENCY", dtHdSo.Rows[0]["CURRENCY"]);
        tbHdSo.SetValue("REDEMP_POINT", dtHdSo.Rows[0]["REDEMP_POINT"]);
        tbHdSo.SetValue("CUST_GRP2", dtHdSo.Rows[0]["CUST_GRP2"]);
        tbHdSo.SetValue("INT_REMARKS", dtHdSo.Rows[0]["INT_REMARKS"]);
        tbHdSo.SetValue("PURCH_DATE", dtHdSo.Rows[0]["PURCH_DATE"]);
        function.SetValue("IS_HD_SO", tbHdSo);

        //5)Zisd0001HdSoPartner
        DataTable dtHdSoPartner = dsData.Tables["dtHdSoPartner"];
        IRfcStructure tbHdSoPartner = function.GetStructure("IS_PARTNER");
        //IRfcTable tbHdSoPartner = function.GetTable("IS_PARTNER");
        //tbHdSoPartner.Append();
        tbHdSoPartner.SetValue("SOLD_TO_ID", dtHdSoPartner.Rows[0]["SOLD_TO_ID"]);
        tbHdSoPartner.SetValue("SHIP_TO_ID", dtHdSoPartner.Rows[0]["SHIP_TO_ID"]);
        tbHdSoPartner.SetValue("EMPLOYEE_ID", dtHdSoPartner.Rows[0]["EMPLOYEE_ID"]);
        function.SetValue("IS_PARTNER", tbHdSoPartner);


        //6)Zisd0001ItSo
        DataTable dtItSo = dsData.Tables["dtItSo"];
        IRfcTable tbItSo = function.GetTable("T_IT_SO");
        for (int i = 0; i < dtItSo.Rows.Count; i++)
        {
            tbItSo.Append();
            tbItSo.SetValue("IT_INDICATOR", dtItSo.Rows[i]["IT_INDICATOR"]);
            tbItSo.SetValue("POITM_NO_S", dtItSo.Rows[i]["POITM_NO_S"]);
            tbItSo.SetValue("MATERIAL", dtItSo.Rows[i]["MATERIAL"]);
            tbItSo.SetValue("ITM_TYPE_USAGE", dtItSo.Rows[i]["ITM_TYPE_USAGE"]);
            tbItSo.SetValue("COND_VALUE", dtItSo.Rows[i]["COND_VALUE"]);
            tbItSo.SetValue("REQ_QTY", dtItSo.Rows[i]["REQ_QTY"]);
        }
        function.SetValue("T_IT_SO", tbItSo);

        //Call SAP Function     
        SODoc_CreateDatatable();
        try
        {
            //execute the function
            function.Invoke(rfcDest);
            dtSOException.Rows.Add("Y");
        }
        catch (Exception e)
        {
            dtSOException.Rows.Add(e.Message);
        }

        IRfcTable tbReturn = function.GetTable("T_Return");

        //----dtSODoc----//
        DataRow drSODoc = dtSODoc.NewRow();
        drSODoc["E_SALES_DOC"] = function.GetValue("E_SALES_DOC").ToString();
        drSODoc["E_ADVANCE_REC_DOC"] = function.GetValue("E_ADVANCE_REC_DOC").ToString();
        drSODoc["E_COMP_CODE"] = function.GetValue("E_COMP_CODE").ToString();
        drSODoc["E_FISCAL_YEAR"] = function.GetValue("E_FISCAL_YEAR").ToString();
        dtSODoc.Rows.Add(drSODoc);

        //Create Dataset
        return SODoc_CreateDataset(tbReturn, dtSODoc, dtSOException);
    }
    public void SODoc_CreateDatatable()
    {
        //dtSOException
        dtSOException.Columns.Add("Message");
        dtSOException.Clear();

        //dtSOReturn
        dtSOReturn.Columns.Add("Type"); dtSOReturn.Columns.Add("Id"); dtSOReturn.Columns.Add("Number");
        dtSOReturn.Columns.Add("Message"); dtSOReturn.Columns.Add("Log_No"); dtSOReturn.Columns.Add("Log_Msg_No");
        dtSOReturn.Columns.Add("Message_V1"); dtSOReturn.Columns.Add("Message_V2"); 
        dtSOReturn.Columns.Add("Message_V3"); dtSOReturn.Columns.Add("Message_V4");
        dtSOReturn.Columns.Add("Parameter"); dtSOReturn.Columns.Add("Row"); dtSOReturn.Columns.Add("Field"); dtSOReturn.Columns.Add("SYSTEM");
        dtSOReturn.Clear();

        //dtSODoc
        dtSODoc.Columns.Add("E_SALES_DOC"); dtSODoc.Columns.Add("E_ADVANCE_REC_DOC");
        dtSODoc.Columns.Add("E_COMP_CODE"); dtSODoc.Columns.Add("E_FISCAL_YEAR"); 
        dtSODoc.Clear();

    }
    public DataSet SODoc_CreateDataset(IRfcTable tbReturn, DataTable dtSODoc, DataTable dtException)
    {
        //tbReturn
        for (int i = 0; i < tbReturn.RowCount; i++)
        {
            tbReturn.CurrentIndex = i;
            dtSOReturn.Rows.Add
                (
                tbReturn.GetString("Type"), tbReturn.GetString("Id"), tbReturn.GetString("Number"),
                tbReturn.GetString("Message"), tbReturn.GetString("Log_No"), tbReturn.GetString("Log_Msg_No"),
                tbReturn.GetString("Message_V1"), tbReturn.GetString("Message_V2"),
                tbReturn.GetString("Message_V3"), tbReturn.GetString("Message_V4"),
                tbReturn.GetString("Parameter"), tbReturn.GetString("Row"), tbReturn.GetString("Field"), tbReturn.GetString("SYSTEM")
                );
        }
        
        DataSet dsResult = new DataSet("dsResult");
        dsResult.Tables.Add(dtException);
        dsResult.Tables.Add(dtSOReturn);
        dsResult.Tables.Add(dtSODoc);

        return dsResult;
    }
    #endregion


    #region "Test"
    ////public String test(string statusStr)
    ////{
    ////    DataSet ds = new DataSet("ds");
    ////    ds = CreateSODocTest4();
    ////    return "Yes";
    ////}
    //////Create SO Doc
    ////public DataSet CreateSODocTest4()
    ////{
    ////    RfcConfigParameters rfc = new RfcConfigParameters();
    ////    rfc.Add(RfcConfigParameters.Name, "CreateSODoc");
    ////    rfc.Add(RfcConfigParameters.AppServerHost, "192.168.2.53");
    ////    rfc.Add(RfcConfigParameters.Client, "220");
    ////    rfc.Add(RfcConfigParameters.User, "basis");
    ////    rfc.Add(RfcConfigParameters.Password, "demo2013");
    ////    rfc.Add(RfcConfigParameters.SystemNumber, "00");
    ////    rfc.Add(RfcConfigParameters.Language, "EN");

    ////    RfcDestination rfcDest = RfcDestinationManager.GetDestination(rfc);
    ////    RfcRepository rfcRep = rfcDest.Repository;
    ////    IRfcFunction function = rfcRep.CreateFunction("ZISD0001_POST_SO_DOC");
    ////    // IRfcStructure structInputs = destination.Repository.GetStructureMetadata("ZECOM_VA01").CreateStructure();  // Structure Name

    ////    function["I_TEST"].SetValue("");

    ////    //1)Zisd0001HdSoShipTo
    ////    IRfcStructure tbShipTo = function.GetStructure("IS_ADDR_SHIP_TO");
    ////    tbShipTo.SetValue("NAME", "");
    ////    tbShipTo.SetValue("NAME_2", "");
    ////    tbShipTo.SetValue("NAME_3", "");
    ////    tbShipTo.SetValue("NAME_4", "");
    ////    tbShipTo.SetValue("STREET_LNG", "");
    ////    tbShipTo.SetValue("STR_SUPPL1", "");
    ////    tbShipTo.SetValue("DISTRICT", "");
    ////    tbShipTo.SetValue("CITY", "");
    ////    tbShipTo.SetValue("POSTL_COD1", "");
    ////    tbShipTo.SetValue("COUNTRY", "");
    ////    tbShipTo.SetValue("TEL1_NUMBR", "");
    ////    tbShipTo.SetValue("TEL1_EXT", "");
    ////    tbShipTo.SetValue("FAX_NUMBER", "");
    ////    tbShipTo.SetValue("FAX_EXTENS", "");
    ////    tbShipTo.SetValue("E_MAIL", "");
    ////    tbShipTo.SetValue("CONTACT_PERSON", "");
    ////    tbShipTo.SetValue("TRANSPZONE", "");
    ////    function["IS_ADDR_SHIP_TO"].SetValue(tbShipTo);

    ////    //2)Zisd0001HdSoSoldTo
    ////    IRfcStructure tbSoldTo = function.GetStructure("IS_ADDR_SOLD_TO");
    ////    tbSoldTo.SetValue("NAME", "");
    ////    tbSoldTo.SetValue("NAME_2", "");
    ////    tbSoldTo.SetValue("NAME_3", "");
    ////    tbSoldTo.SetValue("NAME_4", "");
    ////    tbSoldTo.SetValue("STREET_LNG", "");
    ////    tbSoldTo.SetValue("STR_SUPPL1", "");
    ////    tbSoldTo.SetValue("DISTRICT", "");
    ////    tbSoldTo.SetValue("CITY", "");
    ////    tbSoldTo.SetValue("POSTL_COD1", "");
    ////    tbSoldTo.SetValue("COUNTRY", "");
    ////    tbSoldTo.SetValue("TEL1_NUMBR", "");
    ////    tbSoldTo.SetValue("TEL1_EXT", "");
    ////    tbSoldTo.SetValue("FAX_NUMBER", "");
    ////    tbSoldTo.SetValue("FAX_EXTENS", "");
    ////    tbSoldTo.SetValue("E_MAIL", "");
    ////    tbSoldTo.SetValue("CONTACT_PERSON", "");
    ////    tbSoldTo.SetValue("TAX_ID", "");
    ////    tbSoldTo.SetValue("BRANCH_NO", "");
    ////    function["IS_ADDR_SOLD_TO"].SetValue(tbSoldTo);

    ////    //3)Zisd0001HdFi(ส่วนมัดจำ)
    ////    IRfcStructure tbHdFi = function.GetStructure("IS_HD_FI");
    ////    tbHdFi.SetValue("BLDAT", "20131119");
    ////    tbHdFi.SetValue("NEWKO", "1119010");
    ////    tbHdFi.SetValue("FWBAS", "1922");
    ////    tbHdFi.SetValue("WRBTR", "134.54");
    ////    function["IS_HD_FI"].SetValue(tbHdFi);

    ////    //4)Zisd0001HdSo
    ////    //DataTable dtHdSo = dsData.Tables["dtHdSo"];
    ////    //IRfcStructure tbHdSo = function.GetStructure("IS_HD_SO");
    ////    //IRfcTable tbHdSo = function.GetTable("IS_HD_SO");
    ////    //tbHdSo.Append();
    ////    IRfcStructure tbHdSo = function.GetStructure("IS_HD_SO");
    ////    tbHdSo.SetValue("HD_INDICATOR", "01");
    ////    tbHdSo.SetValue("PURCH_NO_S", "47");
    ////    tbHdSo.SetValue("PO_DAT_S", "20131119");
    ////    tbHdSo.SetValue("DOC_TYPE", "Z212");
    ////    tbHdSo.SetValue("SALES_ORG", "103");
    ////    tbHdSo.SetValue("DISTR_CHAN", "20");
    ////    tbHdSo.SetValue("DIVISION", "50");
    ////    tbHdSo.SetValue("REQ_DATE_H", "20131119");
    ////    tbHdSo.SetValue("PURCH_NO_C", "PO_1900000001");
    ////    tbHdSo.SetValue("PLANT", "1035");
    ////    tbHdSo.SetValue("SHIP_COND", "01");
    ////    tbHdSo.SetValue("CURRENCY", "THB");
    ////    var xx = tbHdSo["HD_INDICATOR"].GetValue();
    ////    function["IS_HD_SO"].SetValue(tbHdSo);

    ////    //5)Zisd0001HdSoPartner
    ////    IRfcStructure tbHdSoPartner = function.GetStructure("IS_PARTNER");
    ////    //tbHdSoPartner.SetValue("SOLD_TO_ID", "100333");
    ////    //tbHdSoPartner.SetValue("SHIP_TO_ID", "100333");
    ////    //tbHdSoPartner.SetValue("EMPLOYEE_ID", "1001");
    ////    tbHdSoPartner.SetValue("SOLD_TO_ID", "102367");
    ////    tbHdSoPartner.SetValue("SHIP_TO_ID", "102367");
    ////    tbHdSoPartner.SetValue("EMPLOYEE_ID", "1001");
    ////    //function.SetValue("IS_PARTNER", tbHdSoPartner);
    ////    IRfcTable tbHDSoPartnerImport = rfcDest.Repository.GetStructureMetadata("ZISD0001_HD_SO_PARTNER").CreateTable();
    ////    tbHDSoPartnerImport.Insert(tbHdSoPartner);
    ////    function["IS_PARTNER"].SetValue(tbHdSoPartner);


    ////    //6)Zisd0001ItSo
    ////    //ataTable dtItSo = dsData.Tables["dtItSo"];
    ////    IRfcTable TBItSo = function.GetTable("T_IT_SO");
    ////    IRfcStructure tbItSo = TBItSo.Metadata.LineType.CreateStructure();

    ////    tbItSo.SetValue("IT_INDICATOR", "11");
    ////    tbItSo.SetValue("POITM_NO_S", "10");
    ////    tbItSo.SetValue("MATERIAL", "1000000004");
    ////    tbItSo.SetValue("ITM_TYPE_USAGE", "");
    ////    tbItSo.SetValue("COND_VALUE", "1000");
    ////    tbItSo.SetValue("REQ_QTY", "2");
    ////    TBItSo.Append(tbItSo);

    ////    //tbItSo.Append();
    ////    tbItSo.SetValue("IT_INDICATOR", "11");
    ////    tbItSo.SetValue("POITM_NO_S", "20");
    ////    tbItSo.SetValue("MATERIAL", "1000000005");
    ////    tbItSo.SetValue("ITM_TYPE_USAGE", "");
    ////    tbItSo.SetValue("COND_VALUE", "850");
    ////    tbItSo.SetValue("REQ_QTY", "2"); ;
    ////    TBItSo.Append(tbItSo);

    ////    //tbItSo.Append();
    ////    tbItSo.SetValue("IT_INDICATOR", "11");
    ////    tbItSo.SetValue("POITM_NO_S", "30");
    ////    tbItSo.SetValue("MATERIAL", "1000000005");
    ////    tbItSo.SetValue("ITM_TYPE_USAGE", "FREE");
    ////    tbItSo.SetValue("COND_VALUE", "0");
    ////    tbItSo.SetValue("REQ_QTY", "1"); ;
    ////    TBItSo.Append(tbItSo);

    ////    //tbItSo.Append();
    ////    tbItSo.SetValue("IT_INDICATOR", "11");
    ////    tbItSo.SetValue("POITM_NO_S", "40");
    ////    tbItSo.SetValue("MATERIAL", "3000000001");
    ////    tbItSo.SetValue("ITM_TYPE_USAGE", "");
    ////    tbItSo.SetValue("COND_VALUE", "-300");
    ////    tbItSo.SetValue("REQ_QTY", "1"); ;
    ////    TBItSo.Append(tbItSo);

    ////    //tbItSo.Append();
    ////    tbItSo.SetValue("IT_INDICATOR", "11");
    ////    tbItSo.SetValue("POITM_NO_S", "50");
    ////    tbItSo.SetValue("MATERIAL", "3000000011");
    ////    tbItSo.SetValue("ITM_TYPE_USAGE", "");
    ////    tbItSo.SetValue("COND_VALUE", "400");
    ////    tbItSo.SetValue("REQ_QTY", "1"); ;
    ////    TBItSo.Append(tbItSo);

    ////    //function.SetValue("T_IT_SO", tbItSo);

    ////    //Call SAP Function     
    ////    SODoc_CreateDatatable();
    ////    try
    ////    {
    ////        //execute the function
    ////        function.Invoke(rfcDest);
    ////        dtSOException.Rows.Add("Y");
    ////    }
    ////    catch (Exception e)
    ////    {
    ////        dtSOException.Rows.Add(e.Message);
    ////    }

    ////    IRfcTable tbReturn = function.GetTable("T_Return");

    ////    //----dtSODoc----//
    ////    DataRow drSODoc = dtSODoc.NewRow();
    ////    drSODoc["E_SALES_DOC"] = function.GetValue("E_SALES_DOC").ToString();
    ////    drSODoc["E_ADVANCE_REC_DOC"] = function.GetValue("E_ADVANCE_REC_DOC").ToString();
    ////    drSODoc["E_COMP_CODE"] = function.GetValue("E_COMP_CODE").ToString();
    ////    drSODoc["E_FISCAL_YEAR"] = function.GetValue("E_FISCAL_YEAR").ToString();
    ////    dtSODoc.Rows.Add(drSODoc);

    ////    //Create Dataset
    ////    return SODoc_CreateDataset(tbReturn, dtSODoc, dtSOException);
    ////}
    ////public DataSet CreateSODocTest3()
    ////{
    ////    RfcConfigParameters rfc = new RfcConfigParameters();
    ////    rfc.Add(RfcConfigParameters.Name, "CreateSODoc");
    ////    rfc.Add(RfcConfigParameters.AppServerHost, "192.168.2.53");
    ////    rfc.Add(RfcConfigParameters.Client, "220");
    ////    rfc.Add(RfcConfigParameters.User, "basis");
    ////    rfc.Add(RfcConfigParameters.Password, "demo2013");
    ////    rfc.Add(RfcConfigParameters.SystemNumber, "00");
    ////    rfc.Add(RfcConfigParameters.Language, "EN");

    ////    RfcDestination rfcDest = RfcDestinationManager.GetDestination(rfc);
    ////    RfcRepository rfcRep = rfcDest.Repository;
    ////    IRfcFunction function = rfcRep.CreateFunction("ZISD0001_POST_SO_DOC");
    ////    // IRfcStructure structInputs = destination.Repository.GetStructureMetadata("ZECOM_VA01").CreateStructure();  // Structure Name

    ////    //1)Zisd0001HdSoShipTo
    ////    IRfcStructure tbShipTo = rfcDest.Repository.GetStructureMetadata("ZISD0001_HD_SO_SHIP_TO").CreateStructure();
    ////    tbShipTo.SetValue("NAME", "");
    ////    tbShipTo.SetValue("NAME_2", "");
    ////    tbShipTo.SetValue("NAME_3", "");
    ////    tbShipTo.SetValue("NAME_4", "");
    ////    tbShipTo.SetValue("STREET_LNG", "");
    ////    tbShipTo.SetValue("STR_SUPPL1", "");
    ////    tbShipTo.SetValue("DISTRICT", "");
    ////    tbShipTo.SetValue("CITY", "");
    ////    tbShipTo.SetValue("POSTL_COD1", "");
    ////    tbShipTo.SetValue("COUNTRY", "");
    ////    tbShipTo.SetValue("TEL1_NUMBR", "");
    ////    tbShipTo.SetValue("TEL1_EXT", "");
    ////    tbShipTo.SetValue("FAX_NUMBER", "");
    ////    tbShipTo.SetValue("FAX_EXTENS", "");
    ////    tbShipTo.SetValue("E_MAIL", "");
    ////    tbShipTo.SetValue("CONTACT_PERSON", "");
    ////    tbShipTo.SetValue("TRANSPZONE", "");
    ////    function["IS_ADDR_SHIP_TO"].SetValue(tbShipTo);

    ////    //2)Zisd0001HdSoSoldTo
    ////    IRfcStructure tbSoldTo = rfcDest.Repository.GetStructureMetadata("ZISD0001_HD_SO_SOLD_TO").CreateStructure();
    ////    tbSoldTo.SetValue("NAME", "");
    ////    tbSoldTo.SetValue("NAME_2", "");
    ////    tbSoldTo.SetValue("NAME_3", "");
    ////    tbSoldTo.SetValue("NAME_4", "");
    ////    tbSoldTo.SetValue("STREET_LNG", "");
    ////    tbSoldTo.SetValue("STR_SUPPL1", "");
    ////    tbSoldTo.SetValue("DISTRICT", "");
    ////    tbSoldTo.SetValue("CITY", "");
    ////    tbSoldTo.SetValue("POSTL_COD1", "");
    ////    tbSoldTo.SetValue("COUNTRY", "");
    ////    tbSoldTo.SetValue("TEL1_NUMBR", "");
    ////    tbSoldTo.SetValue("TEL1_EXT", "");
    ////    tbSoldTo.SetValue("FAX_NUMBER", "");
    ////    tbSoldTo.SetValue("FAX_EXTENS", "");
    ////    tbSoldTo.SetValue("E_MAIL", "");
    ////    tbSoldTo.SetValue("CONTACT_PERSON", "");
    ////    tbSoldTo.SetValue("TAX_ID", "");
    ////    tbSoldTo.SetValue("BRANCH_NO", "");
    ////    function["IS_ADDR_SOLD_TO"].SetValue(tbSoldTo);

    ////    //3)Zisd0001HdFi(ส่วนมัดจำ)
    ////    IRfcStructure tbHdFi = rfcDest.Repository.GetStructureMetadata("ZISD0001_HD_FI").CreateStructure();
    ////    tbHdFi.SetValue("BLDAT", "20131119");
    ////    tbHdFi.SetValue("NEWKO", "1119010");
    ////    tbHdFi.SetValue("FWBAS", "1922");
    ////    tbHdFi.SetValue("WRBTR", "134.54");
    ////    function["IS_HD_FI"].SetValue(tbHdFi);

    ////    //4)Zisd0001HdSo
    ////    //DataTable dtHdSo = dsData.Tables["dtHdSo"];
    ////    //IRfcStructure tbHdSo = function.GetStructure("IS_HD_SO");
    ////    //IRfcTable tbHdSo = function.GetTable("IS_HD_SO");
    ////    //tbHdSo.Append();
    ////    IRfcStructure tbHdSo = rfcDest.Repository.GetStructureMetadata("ZISD0001_HD_SO").CreateStructure();
    ////    tbHdSo.SetValue("HD_INDICATOR", "01");
    ////    tbHdSo.SetValue("PURCH_NO_S", "WO_1900000001");
    ////    tbHdSo.SetValue("PO_DAT_S", "20131119");
    ////    tbHdSo.SetValue("DOC_TYPE", "Z212");
    ////    tbHdSo.SetValue("SALES_ORG", "103");
    ////    tbHdSo.SetValue("DISTR_CHAN", "20");
    ////    tbHdSo.SetValue("DIVISION", "50");
    ////    tbHdSo.SetValue("REQ_DATE_H", "20131119");
    ////    tbHdSo.SetValue("PURCH_NO_C", "PO_1900000001");
    ////    tbHdSo.SetValue("PLANT", "1035");
    ////    tbHdSo.SetValue("SHIP_COND", "01");
    ////    tbHdSo.SetValue("CURRENCY", "THB");
    ////    var xx = tbHdSo["HD_INDICATOR"].GetValue();
    ////    function["IS_HD_SO"].SetValue(tbHdSo);

    ////    //5)Zisd0001HdSoPartner
    ////    IRfcStructure tbHdSoPartner = rfcDest.Repository.GetStructureMetadata("ZISD0001_HD_SO_PARTNER").CreateStructure();
    ////    //tbHdSoPartner.SetValue("SOLD_TO_ID", "100333");
    ////    //tbHdSoPartner.SetValue("SHIP_TO_ID", "100333");
    ////    //tbHdSoPartner.SetValue("EMPLOYEE_ID", "1001");
    ////    tbHdSoPartner.SetValue("SOLD_TO_ID", "102367");
    ////    tbHdSoPartner.SetValue("SHIP_TO_ID", "102367");
    ////    tbHdSoPartner.SetValue("EMPLOYEE_ID", "1001");
    ////    //function.SetValue("IS_PARTNER", tbHdSoPartner);
    ////    IRfcTable tbHDSoPartnerImport = rfcDest.Repository.GetStructureMetadata("ZISD0001_HD_SO_PARTNER").CreateTable();
    ////    tbHDSoPartnerImport.Insert(tbHdSoPartner);
    ////    function["IS_PARTNER"].SetValue(tbHdSoPartner);


    ////    //6)Zisd0001ItSo
    ////    //ataTable dtItSo = dsData.Tables["dtItSo"];
    ////    IRfcTable TBItSo = function.GetTable("T_IT_SO");
    ////    IRfcStructure tbItSo = TBItSo.Metadata.LineType.CreateStructure();

    ////    tbItSo.SetValue("IT_INDICATOR", "11");
    ////    tbItSo.SetValue("POITM_NO_S", "10");
    ////    tbItSo.SetValue("MATERIAL", "1000000004");
    ////    tbItSo.SetValue("ITM_TYPE_USAGE", "");
    ////    tbItSo.SetValue("COND_VALUE", "1000");
    ////    tbItSo.SetValue("REQ_QTY", "2");
    ////    TBItSo.Append(tbItSo);

    ////    //tbItSo.Append();
    ////    tbItSo.SetValue("IT_INDICATOR", "11");
    ////    tbItSo.SetValue("POITM_NO_S", "20");
    ////    tbItSo.SetValue("MATERIAL", "1000000005");
    ////    tbItSo.SetValue("ITM_TYPE_USAGE", "");
    ////    tbItSo.SetValue("COND_VALUE", "850");
    ////    tbItSo.SetValue("REQ_QTY", "2"); ;
    ////    TBItSo.Append(tbItSo);

    ////    //tbItSo.Append();
    ////    tbItSo.SetValue("IT_INDICATOR", "11");
    ////    tbItSo.SetValue("POITM_NO_S", "30");
    ////    tbItSo.SetValue("MATERIAL", "1000000005");
    ////    tbItSo.SetValue("ITM_TYPE_USAGE", "FREE");
    ////    tbItSo.SetValue("COND_VALUE", "0");
    ////    tbItSo.SetValue("REQ_QTY", "1"); ;
    ////    TBItSo.Append(tbItSo);

    ////    //tbItSo.Append();
    ////    tbItSo.SetValue("IT_INDICATOR", "11");
    ////    tbItSo.SetValue("POITM_NO_S", "40");
    ////    tbItSo.SetValue("MATERIAL", "3000000001");
    ////    tbItSo.SetValue("ITM_TYPE_USAGE", "");
    ////    tbItSo.SetValue("COND_VALUE", "-300");
    ////    tbItSo.SetValue("REQ_QTY", "1"); ;
    ////    TBItSo.Append(tbItSo);

    ////    //tbItSo.Append();
    ////    tbItSo.SetValue("IT_INDICATOR", "11");
    ////    tbItSo.SetValue("POITM_NO_S", "50");
    ////    tbItSo.SetValue("MATERIAL", "3000000011");
    ////    tbItSo.SetValue("ITM_TYPE_USAGE", "");
    ////    tbItSo.SetValue("COND_VALUE", "400");
    ////    tbItSo.SetValue("REQ_QTY", "1"); ;
    ////    TBItSo.Append(tbItSo);

    ////    //function.SetValue("T_IT_SO", tbItSo);

    ////    //Call SAP Function     
    ////    SODoc_CreateDatatable();
    ////    try
    ////    {
    ////        //execute the function
    ////        function.Invoke(rfcDest);
    ////        dtSOException.Rows.Add("Y");
    ////    }
    ////    catch (Exception e)
    ////    {
    ////        dtSOException.Rows.Add(e.Message);
    ////    }

    ////    IRfcTable tbReturn = function.GetTable("T_Return");

    ////    //----dtSODoc----//
    ////    DataRow drSODoc = dtSODoc.NewRow();
    ////    drSODoc["E_SALES_DOC"] = function.GetValue("E_SALES_DOC").ToString();
    ////    drSODoc["E_ADVANCE_REC_DOC"] = function.GetValue("E_ADVANCE_REC_DOC").ToString();
    ////    drSODoc["E_COMP_CODE"] = function.GetValue("E_COMP_CODE").ToString();
    ////    drSODoc["E_FISCAL_YEAR"] = function.GetValue("E_FISCAL_YEAR").ToString();
    ////    dtSODoc.Rows.Add(drSODoc);

    ////    //Create Dataset
    ////    return SODoc_CreateDataset(tbReturn, dtSODoc, dtSOException);
    ////}
    ////public DataSet CreateSODocTest()
    ////{
    ////    RfcConfigParameters rfc = new RfcConfigParameters();
    ////    rfc.Add(RfcConfigParameters.Name, "CreateSODoc");
    ////    rfc.Add(RfcConfigParameters.AppServerHost, "192.168.2.53");
    ////    rfc.Add(RfcConfigParameters.Client, "220");
    ////    rfc.Add(RfcConfigParameters.User, "basis");
    ////    rfc.Add(RfcConfigParameters.Password, "demo2013");
    ////    rfc.Add(RfcConfigParameters.SystemNumber, "00");
    ////    rfc.Add(RfcConfigParameters.Language, "EN");

    ////    RfcDestination rfcDest = RfcDestinationManager.GetDestination(rfc);
    ////    RfcRepository rfcRep = rfcDest.Repository;
    ////    IRfcFunction function = rfcRep.CreateFunction("ZISD0001_POST_SO_DOC");
    ////    // IRfcStructure structInputs = destination.Repository.GetStructureMetadata("ZECOM_VA01").CreateStructure();  // Structure Name

    ////    //1)Zisd0001HdSoShipTo
    ////    //DataTable dtShipTo = dsData.Tables["dtShipTo"];
    ////    //IRfcStructure tbShipTo = function.GetStructure("IS_ADDR_SHIP_TO");
    ////    //IRfcTable tbShipTo = function.GetTable("IS_ADDR_SHIP_TO");
    ////    //tbShipTo.Append(); 
    ////    IRfcStructure tbShipTo = rfcDest.Repository.GetStructureMetadata("ZISD0001_HD_SO_SHIP_TO").CreateStructure();
    ////    tbShipTo.SetValue("NAME", "บริษัท One-time_Name1");
    ////    tbShipTo.SetValue("NAME_2", "บริษัท One-time_Name2");
    ////    tbShipTo.SetValue("NAME_3", "");
    ////    tbShipTo.SetValue("NAME_4", "");
    ////    tbShipTo.SetValue("STREET_LNG", "11/2 ถนนพระราม 4");
    ////    tbShipTo.SetValue("STR_SUPPL1", "");
    ////    tbShipTo.SetValue("DISTRICT", "");
    ////    tbShipTo.SetValue("CITY", "กรุงเทพฯ");
    ////    tbShipTo.SetValue("POSTL_COD1", "10120");
    ////    tbShipTo.SetValue("COUNTRY", "TH");
    ////    tbShipTo.SetValue("TEL1_NUMBR", "");
    ////    tbShipTo.SetValue("TEL1_EXT", "");
    ////    tbShipTo.SetValue("FAX_NUMBER", "");
    ////    tbShipTo.SetValue("FAX_EXTENS", "");
    ////    tbShipTo.SetValue("E_MAIL", "");
    ////    tbShipTo.SetValue("CONTACT_PERSON", "");
    ////    tbShipTo.SetValue("TRANSPZONE", "Z100000028");
    ////    //function.SetValue("IS_ADDR_SHIP_TO", tbShipTo);
    ////    IRfcTable tbShipToImport = rfcDest.Repository.GetStructureMetadata("ZISD0001_HD_SO_SHIP_TO").CreateTable();
    ////    tbShipToImport.Insert(tbShipTo);
    ////    function["IS_ADDR_SHIP_TO"].SetValue(tbShipTo);

    ////    //2)Zisd0001HdSoSoldTo
    ////    //DataTable dtSoldTo = dsData.Tables["dtSoldTo"];
    ////    //IRfcStructure tbSoldTo = function.GetStructure("IS_ADDR_SOLD_TO");
    ////    //IRfcTable tbSoldTo = function.GetTable("IS_ADDR_SOLD_TO");
    ////    //tbSoldTo.Append();
    ////    IRfcStructure tbSoldTo = rfcDest.Repository.GetStructureMetadata("ZISD0001_HD_SO_SOLD_TO").CreateStructure();
    ////    tbSoldTo.SetValue("NAME", "บริษัท One-time_Name1");
    ////    tbSoldTo.SetValue("NAME_2", "บริษัท One-time_Name2");
    ////    tbSoldTo.SetValue("NAME_3", "");
    ////    tbSoldTo.SetValue("NAME_4", "");
    ////    tbSoldTo.SetValue("STREET_LNG", "11/2 ถนนพระราม 4");
    ////    tbSoldTo.SetValue("STR_SUPPL1", "");
    ////    tbSoldTo.SetValue("DISTRICT", "");
    ////    tbSoldTo.SetValue("CITY", "กรุงเทพฯ");
    ////    tbSoldTo.SetValue("POSTL_COD1", "10120");
    ////    tbSoldTo.SetValue("COUNTRY", "TH");
    ////    tbSoldTo.SetValue("TEL1_NUMBR", "");
    ////    tbSoldTo.SetValue("TEL1_EXT", "");
    ////    tbSoldTo.SetValue("FAX_NUMBER", "");
    ////    tbSoldTo.SetValue("FAX_EXTENS", "");
    ////    tbSoldTo.SetValue("E_MAIL", "");
    ////    tbSoldTo.SetValue("CONTACT_PERSON", "");
    ////    tbSoldTo.SetValue("TAX_ID", "0000000000000");
    ////    tbSoldTo.SetValue("BRANCH_NO", "00");
    ////    //function.SetValue("IS_ADDR_SOLD_TO", tbSoldTo);
    ////    IRfcTable tbSoldToImport = rfcDest.Repository.GetStructureMetadata("ZISD0001_HD_SO_SOLD_TO").CreateTable();
    ////    tbSoldToImport.Insert(tbSoldTo);
    ////    function["IS_ADDR_SOLD_TO"].SetValue(tbSoldTo);

    ////    //3)Zisd0001HdFi(ส่วนมัดจำ)
    ////    //DataTable dtHdFi = dsData.Tables["dtHdFi"];
    ////    //IRfcStructure tbHdFi = function.GetStructure("IS_HD_FI");
    ////    IRfcStructure tbHdFi = rfcDest.Repository.GetStructureMetadata("ZISD0001_HD_FI").CreateStructure();
    ////    //IRfcTable tbHdFi = function.GetTable("IS_HD_FI");
    ////    //tbHdFi.Append();
    ////    tbHdFi.SetValue("BLDAT", "");
    ////    tbHdFi.SetValue("NEWKO", "");
    ////    tbHdFi.SetValue("FWBAS", "0");
    ////    tbHdFi.SetValue("WRBTR", "0");
    ////    //function.SetValue("IS_HD_FI", tbHdFi);
    ////    IRfcTable tbHDFIImport = rfcDest.Repository.GetStructureMetadata("ZISD0001_HD_FI").CreateTable();
    ////    tbHDFIImport.Insert(tbHdFi);
    ////    function["IS_HD_FI"].SetValue(tbHdFi);

    ////    //4)Zisd0001HdSo
    ////    //DataTable dtHdSo = dsData.Tables["dtHdSo"];
    ////    //IRfcStructure tbHdSo = function.GetStructure("IS_HD_SO");
    ////    //IRfcTable tbHdSo = function.GetTable("IS_HD_SO");
    ////    //tbHdSo.Append();
    ////    IRfcStructure tbHdSo = rfcDest.Repository.GetStructureMetadata("ZISD0001_HD_SO").CreateStructure();
    ////    tbHdSo.SetValue("HD_INDICATOR", "01");
    ////    tbHdSo.SetValue("PURCH_NO_S", "45");
    ////    tbHdSo.SetValue("PO_DAT_S", "20131204");
    ////    tbHdSo.SetValue("DOC_TYPE", "Z211");
    ////    tbHdSo.SetValue("SALES_ORG", "103");
    ////    tbHdSo.SetValue("DISTR_CHAN", "20");
    ////    tbHdSo.SetValue("DIVISION", "50");
    ////    tbHdSo.SetValue("REQ_DATE_H", "20131206");
    ////    tbHdSo.SetValue("PURCH_NO_C", ".");
    ////    tbHdSo.SetValue("PLANT", "1035");
    ////    tbHdSo.SetValue("SHIP_COND", "01");
    ////    tbHdSo.SetValue("CURRENCY", "THB");
    ////    //function.SetValue("IS_HD_SO", tbHdSo);
    ////    IRfcTable tbHDSoImport = rfcDest.Repository.GetStructureMetadata("ZISD0001_HD_SO").CreateTable();
    ////    tbHDSoImport.Insert(tbHdSo);
    ////    function["IS_HD_SO"].SetValue(tbHdSo);
    ////    var xx = tbHdSo["HD_INDICATOR"].GetValue();

    ////    //5)Zisd0001HdSoPartner
    ////    //DataTable dtHdSoPartner = dsData.Tables["dtHdSoPartner"];
    ////    //IRfcStructure tbHdSoPartner = function.GetStructure("IS_PARTNER");
    ////    //IRfcTable tbHdSoPartner = function.GetTable("IS_PARTNER");
    ////    //tbHdSoPartner.Append();
    ////    IRfcStructure tbHdSoPartner = rfcDest.Repository.GetStructureMetadata("ZISD0001_HD_SO_PARTNER").CreateStructure();
    ////    //tbHdSoPartner.SetValue("SOLD_TO_ID", "100333");
    ////    //tbHdSoPartner.SetValue("SHIP_TO_ID", "100333");
    ////    //tbHdSoPartner.SetValue("EMPLOYEE_ID", "1001");
    ////    tbHdSoPartner.SetValue("SOLD_TO_ID", "OT03");
    ////    tbHdSoPartner.SetValue("SHIP_TO_ID", "OT03");
    ////    tbHdSoPartner.SetValue("EMPLOYEE_ID", "1001");
    ////    //function.SetValue("IS_PARTNER", tbHdSoPartner);
    ////    IRfcTable tbHDSoPartnerImport = rfcDest.Repository.GetStructureMetadata("ZISD0001_HD_SO_PARTNER").CreateTable();
    ////    tbHDSoPartnerImport.Insert(tbHdSoPartner);
    ////    function["IS_PARTNER"].SetValue(tbHdSoPartner);


    ////    //6)Zisd0001ItSo
    ////    //ataTable dtItSo = dsData.Tables["dtItSo"];
    ////    IRfcTable TBItSo = function.GetTable("T_IT_SO");
    ////    IRfcStructure tbItSo = TBItSo.Metadata.LineType.CreateStructure();

    ////    tbItSo.SetValue("IT_INDICATOR", "11");
    ////    tbItSo.SetValue("POITM_NO_S", "10");
    ////    tbItSo.SetValue("MATERIAL", "1000000004");
    ////    tbItSo.SetValue("ITM_TYPE_USAGE", "");
    ////    tbItSo.SetValue("COND_VALUE", "1000");
    ////    tbItSo.SetValue("REQ_QTY", "2");
    ////    TBItSo.Append(tbItSo);

    ////    //tbItSo.Append();
    ////    tbItSo.SetValue("IT_INDICATOR", "11");
    ////    tbItSo.SetValue("POITM_NO_S", "20");
    ////    tbItSo.SetValue("MATERIAL", "1000000005");
    ////    tbItSo.SetValue("ITM_TYPE_USAGE", "");
    ////    tbItSo.SetValue("COND_VALUE", "850");
    ////    tbItSo.SetValue("REQ_QTY", "2"); ;
    ////    TBItSo.Append(tbItSo);

    ////    //tbItSo.Append();
    ////    tbItSo.SetValue("IT_INDICATOR", "11");
    ////    tbItSo.SetValue("POITM_NO_S", "30");
    ////    tbItSo.SetValue("MATERIAL", "1000000005");
    ////    tbItSo.SetValue("ITM_TYPE_USAGE", "FREE");
    ////    tbItSo.SetValue("COND_VALUE", "0");
    ////    tbItSo.SetValue("REQ_QTY", "1"); ;
    ////    TBItSo.Append(tbItSo);

    ////    //tbItSo.Append();
    ////    tbItSo.SetValue("IT_INDICATOR", "11");
    ////    tbItSo.SetValue("POITM_NO_S", "40");
    ////    tbItSo.SetValue("MATERIAL", "3000000001");
    ////    tbItSo.SetValue("ITM_TYPE_USAGE", "");
    ////    tbItSo.SetValue("COND_VALUE", "-300");
    ////    tbItSo.SetValue("REQ_QTY", "1"); ;
    ////    TBItSo.Append(tbItSo);

    ////    //tbItSo.Append();
    ////    tbItSo.SetValue("IT_INDICATOR", "11");
    ////    tbItSo.SetValue("POITM_NO_S", "50");
    ////    tbItSo.SetValue("MATERIAL", "3000000011");
    ////    tbItSo.SetValue("ITM_TYPE_USAGE", "");
    ////    tbItSo.SetValue("COND_VALUE", "400");
    ////    tbItSo.SetValue("REQ_QTY", "1"); ;
    ////    TBItSo.Append(tbItSo);

    ////    //function.SetValue("T_IT_SO", tbItSo);

    ////    //Call SAP Function     
    ////    SODoc_CreateDatatable();
    ////    try
    ////    {
    ////        //execute the function
    ////        function.Invoke(rfcDest);
    ////        dtSOException.Rows.Add("Y");
    ////    }
    ////    catch (Exception e)
    ////    {
    ////        dtSOException.Rows.Add(e.Message);
    ////    }

    ////    IRfcTable tbReturn = function.GetTable("T_Return");

    ////    //----dtSODoc----//
    ////    DataRow drSODoc = dtSODoc.NewRow();
    ////    drSODoc["E_SALES_DOC"] = function.GetValue("E_SALES_DOC").ToString();
    ////    drSODoc["E_ADVANCE_REC_DOC"] = function.GetValue("E_ADVANCE_REC_DOC").ToString();
    ////    drSODoc["E_COMP_CODE"] = function.GetValue("E_COMP_CODE").ToString();
    ////    drSODoc["E_FISCAL_YEAR"] = function.GetValue("E_FISCAL_YEAR").ToString();
    ////    dtSODoc.Rows.Add(drSODoc);

    ////    //Create Dataset
    ////    return SODoc_CreateDataset(tbReturn, dtSODoc, dtSOException);
    ////}
    ////public DataSet CreateSODocTest2()
    ////{
    ////    RfcConfigParameters rfc = new RfcConfigParameters();
    ////    rfc.Add(RfcConfigParameters.Name, "CreateSODoc");
    ////    rfc.Add(RfcConfigParameters.AppServerHost, "192.168.2.53");
    ////    rfc.Add(RfcConfigParameters.Client, "220");
    ////    rfc.Add(RfcConfigParameters.User, "basis");
    ////    rfc.Add(RfcConfigParameters.Password, "demo2013");
    ////    rfc.Add(RfcConfigParameters.SystemNumber, "00");
    ////    rfc.Add(RfcConfigParameters.Language, "EN");

    ////    RfcDestination rfcDest = RfcDestinationManager.GetDestination(rfc);
    ////    RfcRepository rfcRep = rfcDest.Repository;
    ////    IRfcFunction function = rfcRep.CreateFunction("ZISD0001_POST_SO_DOC");
    ////    // IRfcStructure structInputs = destination.Repository.GetStructureMetadata("ZECOM_VA01").CreateStructure();  // Structure Name
        
    ////    //1)Zisd0001HdSoShipTo
    ////    //IRfcStructure tbShipTo = rfcDest.Repository.GetStructureMetadata("IS_ADDR_SHIP_TO").CreateStructure();
    ////    //IRfcTable tbShipTo = rfcDest.Repository.GetStructureMetadata("ZISD0001_HD_SO_SHIP_TO").CreateTable();
    ////    //tbShipTo.Insert();
    ////    //tbShipTo.CurrentIndex = tbShipTo.Count - 1; 
    ////    //IRfcTable t_items = function.GetTable("ZISD0001_HD_SO_SHIP_TO");
    ////    RfcStructureMetadata ShipTo = rfcRep.GetStructureMetadata("ZISD0001_HD_SO_SHIP_TO");
    ////    IRfcStructure tbShipTo = ShipTo.CreateStructure();
    ////    tbShipTo.SetValue("NAME", "บริษัท One-time_Name1");
    ////    tbShipTo.SetValue("NAME_2", "บริษัท One-time_Name2");
    ////    tbShipTo.SetValue("NAME_3", "");
    ////    tbShipTo.SetValue("NAME_4", "");
    ////    tbShipTo.SetValue("STREET_LNG", "11/2 ถนนพระราม 4");
    ////    tbShipTo.SetValue("STR_SUPPL1", "");
    ////    tbShipTo.SetValue("DISTRICT", "");
    ////    tbShipTo.SetValue("CITY", "กรุงเทพฯ");
    ////    tbShipTo.SetValue("POSTL_COD1", "10120");
    ////    tbShipTo.SetValue("COUNTRY", "TH");
    ////    tbShipTo.SetValue("TEL1_NUMBR", "");
    ////    tbShipTo.SetValue("TEL1_EXT", "");
    ////    tbShipTo.SetValue("FAX_NUMBER", "");
    ////    tbShipTo.SetValue("FAX_EXTENS", "");
    ////    tbShipTo.SetValue("E_MAIL", "");
    ////    tbShipTo.SetValue("CONTACT_PERSON", "");
    ////    tbShipTo.SetValue("TRANSPZONE", "Z100000028");
    ////    function["IS_ADDR_SHIP_TO"].SetValue(tbShipTo);

    ////    //2)Zisd0001HdSoSoldTo
    ////    //IRfcTable tbSoldTo = function.GetTable("IS_ADDR_SOLD_TO");
    ////    //IRfcTable tbSoldTo = rfcDest.Repository.GetStructureMetadata("ZISD0001_HD_SO_SOLD_TO").CreateTable();
    ////    //tbSoldTo.Insert();
    ////    //tbSoldTo.CurrentIndex = tbSoldTo.Count - 1;
    ////    RfcStructureMetadata SoldTo = rfcRep.GetStructureMetadata("ZISD0001_HD_SO_SOLD_TO");
    ////    IRfcStructure tbSoldTo = SoldTo.CreateStructure();
    ////    tbSoldTo.SetValue("NAME", "บริษัท One-time_Name1");
    ////    tbSoldTo.SetValue("NAME_2", "บริษัท One-time_Name2");
    ////    tbSoldTo.SetValue("NAME_3", "");
    ////    tbSoldTo.SetValue("NAME_4", "");
    ////    tbSoldTo.SetValue("STREET_LNG", "11/2 ถนนพระราม 4");
    ////    tbSoldTo.SetValue("STR_SUPPL1", "");
    ////    tbSoldTo.SetValue("DISTRICT", "");
    ////    tbSoldTo.SetValue("CITY", "กรุงเทพฯ");
    ////    tbSoldTo.SetValue("POSTL_COD1", "10120");
    ////    tbSoldTo.SetValue("COUNTRY", "TH");
    ////    tbSoldTo.SetValue("TEL1_NUMBR", "");
    ////    tbSoldTo.SetValue("TEL1_EXT", "");
    ////    tbSoldTo.SetValue("FAX_NUMBER", "");
    ////    tbSoldTo.SetValue("FAX_EXTENS", "");
    ////    tbSoldTo.SetValue("E_MAIL", "");
    ////    tbSoldTo.SetValue("CONTACT_PERSON", "");
    ////    tbSoldTo.SetValue("TAX_ID", "0000000000000");
    ////    tbSoldTo.SetValue("BRANCH_NO", "00");
    ////    function["IS_ADDR_SOLD_TO"].SetValue(tbSoldTo);


    ////    //3)Zisd0001HdFi(ส่วนมัดจำ)
    ////    //IRfcTable tbHdFi = function.GetTable("IS_HD_FI");
    ////    //IRfcTable tbHdFi = rfcDest.Repository.GetStructureMetadata("ZISD0001_HD_FI").CreateTable();
    ////    //tbHdFi.Insert();
    ////    //tbHdFi.CurrentIndex = tbHdFi.Count - 1;
    ////    RfcStructureMetadata HdFi = rfcRep.GetStructureMetadata("ZISD0001_HD_FI");
    ////    IRfcStructure tbHdFi = HdFi.CreateStructure();
    ////    tbHdFi.SetValue("BLDAT", "");
    ////    tbHdFi.SetValue("NEWKO", "");
    ////    tbHdFi.SetValue("FWBAS", "0");
    ////    tbHdFi.SetValue("WRBTR", "0");
    ////    function["IS_HD_FI"].SetValue(tbHdFi);

    ////    //4)Zisd0001HdSo
    ////    //IRfcTable tbHdSo = function.GetTable("IS_HD_SO");
    ////    //IRfcTable tbHdSo = rfcDest.Repository.GetStructureMetadata("ZISD0001_HD_SO").CreateTable();
    ////    //tbHdSo.Insert();
    ////    //tbHdSo.CurrentIndex = tbHdSo.Count - 1;
    ////    RfcStructureMetadata HdSo = rfcRep.GetStructureMetadata("ZISD0001_HD_SO");
    ////    IRfcStructure tbHdSo = HdSo.CreateStructure();
    ////    tbHdSo.SetValue("HD_INDICATOR", "01");
    ////    tbHdSo.SetValue("PURCH_NO_S", "45");
    ////    tbHdSo.SetValue("PO_DAT_S", "20131204");
    ////    tbHdSo.SetValue("DOC_TYPE", "Z211");
    ////    tbHdSo.SetValue("SALES_ORG", "103");
    ////    tbHdSo.SetValue("DISTR_CHAN", "20");
    ////    tbHdSo.SetValue("DIVISION", "50");
    ////    tbHdSo.SetValue("REQ_DATE_H", "20131206");
    ////    tbHdSo.SetValue("PURCH_NO_C", ".");
    ////    tbHdSo.SetValue("PLANT", "1035");
    ////    tbHdSo.SetValue("SHIP_COND", "01");
    ////    tbHdSo.SetValue("CURRENCY", "THB");
    ////    function["IS_HD_SO"].SetValue(tbHdSo);

    ////    //5)Zisd0001HdSoPartner
    ////    //IRfcTable tbHdSoPartner = function.GetTable("IS_PARTNER");
    ////    //IRfcTable tbHdSoPartner = rfcDest.Repository.GetStructureMetadata("ZISD0001_HD_SO_PARTNER").CreateTable();
    ////    //tbHdSoPartner.Insert();
    ////    //tbHdSoPartner.CurrentIndex = tbHdSoPartner.Count - 1;
    ////    RfcStructureMetadata HdSoPartner = rfcRep.GetStructureMetadata("ZISD0001_HD_SO_PARTNER");
    ////    IRfcStructure tbHdSoPartner = HdSoPartner.CreateStructure();
    ////    tbHdSoPartner.SetValue("SOLD_TO_ID", "OT03");
    ////    tbHdSoPartner.SetValue("SHIP_TO_ID", "OT03");
    ////    tbHdSoPartner.SetValue("EMPLOYEE_ID", "1001");
    ////    function["IS_PARTNER"].SetValue(tbHdSoPartner);


    ////    //6)Zisd0001ItSo
    ////    //ataTable dtItSo = dsData.Tables["dtItSo"];
    ////    IRfcTable TBItSo = function.GetTable("T_IT_SO");
    ////    IRfcStructure tbItSo = TBItSo.Metadata.LineType.CreateStructure();

    ////    tbItSo.SetValue("IT_INDICATOR", "11");
    ////    tbItSo.SetValue("POITM_NO_S", "10");
    ////    tbItSo.SetValue("MATERIAL", "1000000004");
    ////    tbItSo.SetValue("ITM_TYPE_USAGE", "");
    ////    tbItSo.SetValue("COND_VALUE", "1000");
    ////    tbItSo.SetValue("REQ_QTY", "2");
    ////    TBItSo.Append(tbItSo);

    ////    //tbItSo.Append();
    ////    tbItSo.SetValue("IT_INDICATOR", "11");
    ////    tbItSo.SetValue("POITM_NO_S", "20");
    ////    tbItSo.SetValue("MATERIAL", "1000000005");
    ////    tbItSo.SetValue("ITM_TYPE_USAGE", "");
    ////    tbItSo.SetValue("COND_VALUE", "850");
    ////    tbItSo.SetValue("REQ_QTY", "2"); ;
    ////    TBItSo.Append(tbItSo);

    ////    //tbItSo.Append();
    ////    tbItSo.SetValue("IT_INDICATOR", "11");
    ////    tbItSo.SetValue("POITM_NO_S", "30");
    ////    tbItSo.SetValue("MATERIAL", "1000000005");
    ////    tbItSo.SetValue("ITM_TYPE_USAGE", "FREE");
    ////    tbItSo.SetValue("COND_VALUE", "0");
    ////    tbItSo.SetValue("REQ_QTY", "1"); ;
    ////    TBItSo.Append(tbItSo);

    ////    //tbItSo.Append();
    ////    tbItSo.SetValue("IT_INDICATOR", "11");
    ////    tbItSo.SetValue("POITM_NO_S", "40");
    ////    tbItSo.SetValue("MATERIAL", "3000000001");
    ////    tbItSo.SetValue("ITM_TYPE_USAGE", "");
    ////    tbItSo.SetValue("COND_VALUE", "-300");
    ////    tbItSo.SetValue("REQ_QTY", "1"); ;
    ////    TBItSo.Append(tbItSo);

    ////    //tbItSo.Append();
    ////    tbItSo.SetValue("IT_INDICATOR", "11");
    ////    tbItSo.SetValue("POITM_NO_S", "50");
    ////    tbItSo.SetValue("MATERIAL", "3000000011");
    ////    tbItSo.SetValue("ITM_TYPE_USAGE", "");
    ////    tbItSo.SetValue("COND_VALUE", "400");
    ////    tbItSo.SetValue("REQ_QTY", "1"); ;
    ////    TBItSo.Append(tbItSo);

    ////    //function.SetValue("T_IT_SO", tbItSo);

    ////    //Call SAP Function     
    ////    SODoc_CreateDatatable();
    ////    try
    ////    {
    ////        //execute the function
    ////        function.Invoke(rfcDest);
    ////        dtSOException.Rows.Add("Y");
    ////    }
    ////    catch (Exception e)
    ////    {
    ////        dtSOException.Rows.Add(e.Message);
    ////    }

    ////    IRfcTable tbReturn = function.GetTable("T_Return");

    ////    //----dtSODoc----//
    ////    DataRow drSODoc = dtSODoc.NewRow();
    ////    drSODoc["E_SALES_DOC"] = function.GetValue("E_SALES_DOC").ToString();
    ////    drSODoc["E_ADVANCE_REC_DOC"] = function.GetValue("E_ADVANCE_REC_DOC").ToString();
    ////    drSODoc["E_COMP_CODE"] = function.GetValue("E_COMP_CODE").ToString();
    ////    drSODoc["E_FISCAL_YEAR"] = function.GetValue("E_FISCAL_YEAR").ToString();
    ////    dtSODoc.Rows.Add(drSODoc);

    ////    //Create Dataset
    ////    return SODoc_CreateDataset(tbReturn, dtSODoc, dtSOException);
    ////}
    ////public DataSet CreateSODoc_ForTest2(DataSet dsData)
    ////{
    ////    RfcConfigParameters rfc = new RfcConfigParameters();
    ////    rfc.Add(RfcConfigParameters.Name, "CreateSODoc");
    ////    rfc.Add(RfcConfigParameters.AppServerHost, "192.168.2.53");
    ////    rfc.Add(RfcConfigParameters.Client, "220");
    ////    rfc.Add(RfcConfigParameters.User, "basis");
    ////    rfc.Add(RfcConfigParameters.Password, "demo2013");
    ////    rfc.Add(RfcConfigParameters.SystemNumber, "00");
    ////    rfc.Add(RfcConfigParameters.Language, "EN");

    ////    RfcDestination rfcDest = RfcDestinationManager.GetDestination(rfc);
    ////    RfcRepository rfcRep = rfcDest.Repository;
    ////    IRfcFunction function = rfcRep.CreateFunction("ZISD0001_POST_SO_DOC");

    ////    //1)Zisd0001HdSoShipTo
    ////    DataTable dtShipTo = dsData.Tables["dtShipTo"];
    ////    IRfcStructure tbShipTo = rfcDest.Repository.GetStructureMetadata("ZISD0001_HD_SO_SHIP_TO").CreateStructure();
    ////    tbShipTo.SetValue("NAME", dtShipTo.Rows[0]["NAME"]);
    ////    tbShipTo.SetValue("NAME_2", dtShipTo.Rows[0]["NAME_2"]);
    ////    tbShipTo.SetValue("NAME_3", dtShipTo.Rows[0]["NAME_3"]);
    ////    tbShipTo.SetValue("NAME_4", dtShipTo.Rows[0]["NAME_4"]);
    ////    tbShipTo.SetValue("STREET_LNG", dtShipTo.Rows[0]["STREET_LNG"]);
    ////    tbShipTo.SetValue("STR_SUPPL1", dtShipTo.Rows[0]["STR_SUPPL1"]);
    ////    tbShipTo.SetValue("DISTRICT", dtShipTo.Rows[0]["DISTRICT"]);
    ////    tbShipTo.SetValue("CITY", dtShipTo.Rows[0]["CITY"]);
    ////    tbShipTo.SetValue("POSTL_COD1", dtShipTo.Rows[0]["POSTL_COD1"]);
    ////    tbShipTo.SetValue("COUNTRY", dtShipTo.Rows[0]["COUNTRY"]);
    ////    tbShipTo.SetValue("TEL1_NUMBR", dtShipTo.Rows[0]["TEL1_NUMBR"]);
    ////    tbShipTo.SetValue("TEL1_EXT", dtShipTo.Rows[0]["TEL1_EXT"]);
    ////    tbShipTo.SetValue("FAX_NUMBER", dtShipTo.Rows[0]["FAX_NUMBER"]);
    ////    tbShipTo.SetValue("FAX_EXTENS", dtShipTo.Rows[0]["FAX_EXTENS"]);
    ////    tbShipTo.SetValue("E_MAIL", dtShipTo.Rows[0]["E_MAIL"]);
    ////    tbShipTo.SetValue("CONTACT_PERSON", dtShipTo.Rows[0]["CONTACT_PERSON"]);
    ////    tbShipTo.SetValue("TRANSPZONE", dtShipTo.Rows[0]["TRANSPZONE"]);
    ////    IRfcTable tbShipToImport = rfcDest.Repository.GetStructureMetadata("ZISD0001_HD_SO_SHIP_TO").CreateTable();
    ////    tbShipToImport.Insert(tbShipTo);


    ////    //2)Zisd0001HdSoSoldTo
    ////    DataTable dtSoldTo = dsData.Tables["dtSoldTo"];
    ////    IRfcStructure tbSoldTo = rfcDest.Repository.GetStructureMetadata("ZISD0001_HD_SO_SOLD_TO").CreateStructure();
    ////    tbSoldTo.SetValue("NAME", dtSoldTo.Rows[0]["NAME"]);
    ////    tbSoldTo.SetValue("NAME_2", dtSoldTo.Rows[0]["NAME_2"]);
    ////    tbSoldTo.SetValue("NAME_3", dtSoldTo.Rows[0]["NAME_3"]);
    ////    tbSoldTo.SetValue("NAME_4", dtSoldTo.Rows[0]["NAME_4"]);
    ////    tbSoldTo.SetValue("STREET_LNG", dtSoldTo.Rows[0]["STREET_LNG"]);
    ////    tbSoldTo.SetValue("STR_SUPPL1", dtSoldTo.Rows[0]["STR_SUPPL1"]);
    ////    tbSoldTo.SetValue("DISTRICT", dtSoldTo.Rows[0]["DISTRICT"]);
    ////    tbSoldTo.SetValue("CITY", dtSoldTo.Rows[0]["CITY"]);
    ////    tbSoldTo.SetValue("POSTL_COD1", dtSoldTo.Rows[0]["POSTL_COD1"]);
    ////    tbSoldTo.SetValue("COUNTRY", dtSoldTo.Rows[0]["COUNTRY"]);
    ////    tbSoldTo.SetValue("TEL1_NUMBR", dtSoldTo.Rows[0]["TEL1_NUMBR"]);
    ////    tbSoldTo.SetValue("TEL1_EXT", dtSoldTo.Rows[0]["TEL1_EXT"]);
    ////    tbSoldTo.SetValue("FAX_NUMBER", dtSoldTo.Rows[0]["FAX_NUMBER"]);
    ////    tbSoldTo.SetValue("FAX_EXTENS", dtSoldTo.Rows[0]["FAX_EXTENS"]);
    ////    tbSoldTo.SetValue("E_MAIL", dtSoldTo.Rows[0]["E_MAIL"]);
    ////    tbSoldTo.SetValue("CONTACT_PERSON", dtSoldTo.Rows[0]["CONTACT_PERSON"]);
    ////    tbSoldTo.SetValue("TAX_ID", dtSoldTo.Rows[0]["TAX_ID"]);
    ////    tbSoldTo.SetValue("BRANCH_NO", dtSoldTo.Rows[0]["BRANCH_NO"]);
    ////    IRfcTable tbSoldToImport = rfcDest.Repository.GetStructureMetadata("ZISD0001_HD_SO_SOLD_TO").CreateTable();
    ////    tbSoldToImport.Insert(tbSoldTo);

    ////    //3)Zisd0001HdFi(ส่วนมัดจำ)
    ////    DataTable dtHdFi = dsData.Tables["dtHdFi"];
    ////    IRfcStructure tbHdFi = rfcDest.Repository.GetStructureMetadata("ZISD0001_HD_FI").CreateStructure();
    ////    tbHdFi.SetValue("BLDAT", dtHdFi.Rows[0]["BLDAT"]);
    ////    tbHdFi.SetValue("NEWKO", dtHdFi.Rows[0]["NEWKO"]);
    ////    tbHdFi.SetValue("FWBAS", dtHdFi.Rows[0]["FWBAS"]);
    ////    tbHdFi.SetValue("WRBTR", dtHdFi.Rows[0]["WRBTR"]);
    ////    IRfcTable tbHdFiImport = rfcDest.Repository.GetStructureMetadata("ZISD0001_HD_FI").CreateTable();
    ////    tbHdFiImport.Insert(tbHdFi);

    ////    //4)Zisd0001HdSo
    ////    DataTable dtHdSo = dsData.Tables["dtHdSo"];
    ////    IRfcStructure tbHdSo = rfcDest.Repository.GetStructureMetadata("ZISD0001_HD_SO").CreateStructure();
    ////    tbHdSo.SetValue("HD_INDICATOR", dtHdSo.Rows[0]["HD_INDICATOR"]);
    ////    tbHdSo.SetValue("PURCH_NO_S", dtHdSo.Rows[0]["PURCH_NO_S"]);
    ////    tbHdSo.SetValue("PO_DAT_S", dtHdSo.Rows[0]["PO_DAT_S"]);
    ////    tbHdSo.SetValue("DOC_TYPE", dtHdSo.Rows[0]["DOC_TYPE"]);
    ////    tbHdSo.SetValue("SALES_ORG", dtHdSo.Rows[0]["SALES_ORG"]);
    ////    tbHdSo.SetValue("DISTR_CHAN", dtHdSo.Rows[0]["DISTR_CHAN"]);
    ////    tbHdSo.SetValue("DIVISION", dtHdSo.Rows[0]["DIVISION"]);
    ////    tbHdSo.SetValue("REQ_DATE_H", dtHdSo.Rows[0]["REQ_DATE_H"]);
    ////    tbHdSo.SetValue("PURCH_NO_C", dtHdSo.Rows[0]["PURCH_NO_C"]);
    ////    tbHdSo.SetValue("PLANT", dtHdSo.Rows[0]["PLANT"]);
    ////    tbHdSo.SetValue("SHIP_COND", dtHdSo.Rows[0]["SHIP_COND"]);
    ////    tbHdSo.SetValue("CURRENCY", dtHdSo.Rows[0]["CURRENCY"]);
    ////    IRfcTable tbHdSoImport = rfcDest.Repository.GetStructureMetadata("ZISD0001_HD_SO").CreateTable();
    ////    tbHdSoImport.Insert(tbHdSo);

    ////    //5)Zisd0001HdSoPartner
    ////    DataTable dtHdSoPartner = dsData.Tables["dtHdSoPartner"];
    ////    IRfcStructure tbHdSoPartner = rfcDest.Repository.GetStructureMetadata("ZISD0001_HD_SO_PARTNER").CreateStructure();
    ////    tbHdSoPartner.SetValue("SOLD_TO_ID", dtHdSoPartner.Rows[0]["SOLD_TO_ID"]);
    ////    tbHdSoPartner.SetValue("SHIP_TO_ID", dtHdSoPartner.Rows[0]["SHIP_TO_ID"]);
    ////    tbHdSoPartner.SetValue("EMPLOYEE_ID", dtHdSoPartner.Rows[0]["EMPLOYEE_ID"]);
    ////    IRfcTable tbHdSoPartnerImport = rfcDest.Repository.GetStructureMetadata("ZISD0001_HD_SO_PARTNER").CreateTable();
    ////    tbHdSoPartnerImport.Insert(tbHdSoPartner);

    ////    //6)Zisd0001ItSo
    ////    DataTable dtItSo = dsData.Tables["dtItSo"];
    ////    IRfcTable tbItSo = function.GetTable("T_IT_SO");
    ////    for (int i = 0; i < dtItSo.Rows.Count; i++)
    ////    {
    ////        tbItSo.Append();
    ////        tbItSo.SetValue("IT_INDICATOR", dtItSo.Rows[i]["IT_INDICATOR"]);
    ////        tbItSo.SetValue("POITM_NO_S", dtItSo.Rows[i]["POITM_NO_S"]);
    ////        tbItSo.SetValue("MATERIAL", dtItSo.Rows[i]["MATERIAL"]);
    ////        tbItSo.SetValue("ITM_TYPE_USAGE", dtItSo.Rows[i]["ITM_TYPE_USAGE"]);
    ////        tbItSo.SetValue("COND_VALUE", dtItSo.Rows[i]["COND_VALUE"]);
    ////        tbItSo.SetValue("REQ_QTY", dtItSo.Rows[i]["REQ_QTY"]);
    ////    }
    ////    function.SetValue("T_IT_SO", tbItSo);

    ////    //Call SAP Function     
    ////    SODoc_CreateDatatable();
    ////    try
    ////    {
    ////        //execute the function
    ////        function.Invoke(rfcDest);
    ////        dtSOException.Rows.Add("Y");
    ////    }
    ////    catch (Exception e)
    ////    {
    ////        dtSOException.Rows.Add(e.Message);
    ////    }

    ////    IRfcTable tbReturn = function.GetTable("T_Return");

    ////    //----dtSODoc----//
    ////    DataRow drSODoc = dtSODoc.NewRow();
    ////    drSODoc["E_SALES_DOC"] = function.GetValue("E_SALES_DOC").ToString();
    ////    drSODoc["E_ADVANCE_REC_DOC"] = function.GetValue("E_ADVANCE_REC_DOC").ToString();
    ////    drSODoc["E_COMP_CODE"] = function.GetValue("E_COMP_CODE").ToString();
    ////    drSODoc["E_FISCAL_YEAR"] = function.GetValue("E_FISCAL_YEAR").ToString();
    ////    dtSODoc.Rows.Add(drSODoc);

    ////    //Create Dataset
    ////    return SODoc_CreateDataset(tbReturn, dtSODoc, dtSOException);
    ////}
    ////public DataSet CreateSODoc_ForTest(DataSet dsData)
    ////{
    ////    RfcConfigParameters rfc = new RfcConfigParameters();
    ////    rfc.Add(RfcConfigParameters.Name, "CreateSODoc");
    ////    rfc.Add(RfcConfigParameters.AppServerHost, "192.168.2.53");
    ////    rfc.Add(RfcConfigParameters.Client, "220");
    ////    rfc.Add(RfcConfigParameters.User, "basis");
    ////    rfc.Add(RfcConfigParameters.Password, "demo2013");
    ////    rfc.Add(RfcConfigParameters.SystemNumber, "00");
    ////    rfc.Add(RfcConfigParameters.Language, "EN");

    ////    RfcDestination rfcDest = RfcDestinationManager.GetDestination(rfc);
    ////    RfcRepository rfcRep = rfcDest.Repository;
    ////    IRfcFunction function = rfcRep.CreateFunction("ZISD0001_POST_SO_DOC");

    ////    //1)Zisd0001HdSoShipTo
    ////    DataTable dtShipTo = dsData.Tables["dtShipTo"];
    ////    IRfcStructure tbShipTo = function.GetStructure("IS_ADDR_SHIP_TO");
    ////    //IRfcTable tbShipTo = function.GetTable("IS_ADDR_SHIP_TO");
    ////    //tbShipTo.Append();      
    ////    tbShipTo.SetValue("NAME", dtShipTo.Rows[0]["NAME"]);
    ////    tbShipTo.SetValue("NAME_2", dtShipTo.Rows[0]["NAME_2"]);
    ////    tbShipTo.SetValue("NAME_3", dtShipTo.Rows[0]["NAME_3"]);
    ////    tbShipTo.SetValue("NAME_4", dtShipTo.Rows[0]["NAME_4"]);
    ////    tbShipTo.SetValue("STREET_LNG", dtShipTo.Rows[0]["STREET_LNG"]);
    ////    tbShipTo.SetValue("STR_SUPPL1", dtShipTo.Rows[0]["STR_SUPPL1"]);
    ////    tbShipTo.SetValue("DISTRICT", dtShipTo.Rows[0]["DISTRICT"]);
    ////    tbShipTo.SetValue("CITY", dtShipTo.Rows[0]["CITY"]);
    ////    tbShipTo.SetValue("POSTL_COD1", dtShipTo.Rows[0]["POSTL_COD1"]);
    ////    tbShipTo.SetValue("COUNTRY", dtShipTo.Rows[0]["COUNTRY"]);
    ////    tbShipTo.SetValue("TEL1_NUMBR", dtShipTo.Rows[0]["TEL1_NUMBR"]);
    ////    tbShipTo.SetValue("TEL1_EXT", dtShipTo.Rows[0]["TEL1_EXT"]);
    ////    tbShipTo.SetValue("FAX_NUMBER", dtShipTo.Rows[0]["FAX_NUMBER"]);
    ////    tbShipTo.SetValue("FAX_EXTENS", dtShipTo.Rows[0]["FAX_EXTENS"]);
    ////    tbShipTo.SetValue("E_MAIL", dtShipTo.Rows[0]["E_MAIL"]);
    ////    tbShipTo.SetValue("CONTACT_PERSON", dtShipTo.Rows[0]["CONTACT_PERSON"]);
    ////    tbShipTo.SetValue("TRANSPZONE", dtShipTo.Rows[0]["TRANSPZONE"]);
    ////    function.SetValue("IS_ADDR_SHIP_TO", tbShipTo);

    ////    //2)Zisd0001HdSoSoldTo
    ////    DataTable dtSoldTo = dsData.Tables["dtSoldTo"];
    ////    IRfcStructure tbSoldTo = function.GetStructure("IS_ADDR_SOLD_TO");
    ////    //IRfcTable tbSoldTo = function.GetTable("IS_ADDR_SOLD_TO");
    ////    //tbSoldTo.Append();
    ////    tbSoldTo.SetValue("NAME", dtSoldTo.Rows[0]["NAME"]);
    ////    tbSoldTo.SetValue("NAME_2", dtSoldTo.Rows[0]["NAME_2"]);
    ////    tbSoldTo.SetValue("NAME_3", dtSoldTo.Rows[0]["NAME_3"]);
    ////    tbSoldTo.SetValue("NAME_4", dtSoldTo.Rows[0]["NAME_4"]);
    ////    tbSoldTo.SetValue("STREET_LNG", dtSoldTo.Rows[0]["STREET_LNG"]);
    ////    tbSoldTo.SetValue("STR_SUPPL1", dtSoldTo.Rows[0]["STR_SUPPL1"]);
    ////    tbSoldTo.SetValue("DISTRICT", dtSoldTo.Rows[0]["DISTRICT"]);
    ////    tbSoldTo.SetValue("CITY", dtSoldTo.Rows[0]["CITY"]);
    ////    tbSoldTo.SetValue("POSTL_COD1", dtSoldTo.Rows[0]["POSTL_COD1"]);
    ////    tbSoldTo.SetValue("COUNTRY", dtSoldTo.Rows[0]["COUNTRY"]);
    ////    tbSoldTo.SetValue("TEL1_NUMBR", dtSoldTo.Rows[0]["TEL1_NUMBR"]);
    ////    tbSoldTo.SetValue("TEL1_EXT", dtSoldTo.Rows[0]["TEL1_EXT"]);
    ////    tbSoldTo.SetValue("FAX_NUMBER", dtSoldTo.Rows[0]["FAX_NUMBER"]);
    ////    tbSoldTo.SetValue("FAX_EXTENS", dtSoldTo.Rows[0]["FAX_EXTENS"]);
    ////    tbSoldTo.SetValue("E_MAIL", dtSoldTo.Rows[0]["E_MAIL"]);
    ////    tbSoldTo.SetValue("CONTACT_PERSON", dtSoldTo.Rows[0]["CONTACT_PERSON"]);
    ////    tbSoldTo.SetValue("TAX_ID", dtSoldTo.Rows[0]["TAX_ID"]);
    ////    tbSoldTo.SetValue("BRANCH_NO", dtSoldTo.Rows[0]["BRANCH_NO"]);
    ////    function.SetValue("IS_ADDR_SOLD_TO", tbSoldTo);

    ////    //3)Zisd0001HdFi(ส่วนมัดจำ)
    ////    DataTable dtHdFi = dsData.Tables["dtHdFi"];
    ////    IRfcStructure tbHdFi = function.GetStructure("IS_HD_FI");
    ////    //IRfcTable tbHdFi = function.GetTable("IS_HD_FI");
    ////    //tbHdFi.Append();
    ////    tbHdFi.SetValue("BLDAT", dtHdFi.Rows[0]["BLDAT"]);
    ////    tbHdFi.SetValue("NEWKO", dtHdFi.Rows[0]["NEWKO"]);
    ////    tbHdFi.SetValue("FWBAS", dtHdFi.Rows[0]["FWBAS"]);
    ////    tbHdFi.SetValue("WRBTR", dtHdFi.Rows[0]["WRBTR"]);
    ////    function.SetValue("IS_HD_FI", tbHdFi);

    ////    //4)Zisd0001HdSo
    ////    DataTable dtHdSo = dsData.Tables["dtHdSo"];
    ////    IRfcStructure tbHdSo = function.GetStructure("IS_HD_SO");
    ////    //IRfcTable tbHdSo = function.GetTable("IS_HD_SO");
    ////    //tbHdSo.Append();
    ////    tbHdSo.SetValue("HD_INDICATOR", dtHdSo.Rows[0]["HD_INDICATOR"]);
    ////    tbHdSo.SetValue("PURCH_NO_S", dtHdSo.Rows[0]["PURCH_NO_S"]);
    ////    tbHdSo.SetValue("PO_DAT_S", dtHdSo.Rows[0]["PO_DAT_S"]);
    ////    tbHdSo.SetValue("DOC_TYPE", dtHdSo.Rows[0]["DOC_TYPE"]);
    ////    tbHdSo.SetValue("SALES_ORG", dtHdSo.Rows[0]["SALES_ORG"]);
    ////    tbHdSo.SetValue("DISTR_CHAN", dtHdSo.Rows[0]["DISTR_CHAN"]);
    ////    tbHdSo.SetValue("DIVISION", dtHdSo.Rows[0]["DIVISION"]);
    ////    tbHdSo.SetValue("REQ_DATE_H", dtHdSo.Rows[0]["REQ_DATE_H"]);
    ////    tbHdSo.SetValue("PURCH_NO_C", dtHdSo.Rows[0]["PURCH_NO_C"]);
    ////    tbHdSo.SetValue("PLANT", dtHdSo.Rows[0]["PLANT"]);
    ////    tbHdSo.SetValue("SHIP_COND", dtHdSo.Rows[0]["SHIP_COND"]);
    ////    tbHdSo.SetValue("CURRENCY", dtHdSo.Rows[0]["CURRENCY"]);
    ////    function.SetValue("IS_HD_SO", tbHdSo);

    ////    //5)Zisd0001HdSoPartner
    ////    DataTable dtHdSoPartner = dsData.Tables["dtHdSoPartner"];
    ////    IRfcStructure tbHdSoPartner = function.GetStructure("IS_PARTNER");
    ////    //IRfcTable tbHdSoPartner = function.GetTable("IS_PARTNER");
    ////    //tbHdSoPartner.Append();
    ////    tbHdSoPartner.SetValue("SOLD_TO_ID", dtHdSoPartner.Rows[0]["SOLD_TO_ID"]);
    ////    tbHdSoPartner.SetValue("SHIP_TO_ID", dtHdSoPartner.Rows[0]["SHIP_TO_ID"]);
    ////    tbHdSoPartner.SetValue("EMPLOYEE_ID", dtHdSoPartner.Rows[0]["EMPLOYEE_ID"]);
    ////    function.SetValue("IS_PARTNER", tbHdSoPartner);


    ////    //6)Zisd0001ItSo
    ////    DataTable dtItSo = dsData.Tables["dtItSo"];
    ////    IRfcTable tbItSo = function.GetTable("T_IT_SO");
    ////    for (int i = 0; i < dtItSo.Rows.Count; i++)
    ////    {
    ////        tbItSo.Append();
    ////        tbItSo.SetValue("IT_INDICATOR", dtItSo.Rows[i]["IT_INDICATOR"]);
    ////        tbItSo.SetValue("POITM_NO_S", dtItSo.Rows[i]["POITM_NO_S"]);
    ////        tbItSo.SetValue("MATERIAL", dtItSo.Rows[i]["MATERIAL"]);
    ////        tbItSo.SetValue("ITM_TYPE_USAGE", dtItSo.Rows[i]["ITM_TYPE_USAGE"]);
    ////        tbItSo.SetValue("COND_VALUE", dtItSo.Rows[i]["COND_VALUE"]);
    ////        tbItSo.SetValue("REQ_QTY", dtItSo.Rows[i]["REQ_QTY"]);
    ////    }
    ////    function.SetValue("T_IT_SO", tbItSo);

    ////    //Call SAP Function     
    ////    SODoc_CreateDatatable();
    ////    try
    ////    {
    ////        //execute the function
    ////        function.Invoke(rfcDest);
    ////        dtSOException.Rows.Add("Y");
    ////    }
    ////    catch (Exception e)
    ////    {
    ////        dtSOException.Rows.Add(e.Message);
    ////    }

    ////    IRfcTable tbReturn = function.GetTable("T_Return");

    ////    //----dtSODoc----//
    ////    DataRow drSODoc = dtSODoc.NewRow();
    ////    drSODoc["E_SALES_DOC"] = function.GetValue("E_SALES_DOC").ToString();
    ////    drSODoc["E_ADVANCE_REC_DOC"] = function.GetValue("E_ADVANCE_REC_DOC").ToString();
    ////    drSODoc["E_COMP_CODE"] = function.GetValue("E_COMP_CODE").ToString();
    ////    drSODoc["E_FISCAL_YEAR"] = function.GetValue("E_FISCAL_YEAR").ToString();
    ////    dtSODoc.Rows.Add(drSODoc);

    ////    //Create Dataset
    ////    return SODoc_CreateDataset(tbReturn, dtSODoc, dtSOException);
    ////}    
    #endregion

    #region "BackUp"
    //////20160801:ไม่มี RedemtionPoint, CustomerGroup2, InternalRemark
    ////public DataSet CreateSODoc(DataSet dsData)
    ////{
    ////    RfcConfigParameters rfc = new RfcConfigParameters();
    ////    rfc.Add(RfcConfigParameters.Name, "CheckStock");
    ////    rfc.Add(RfcConfigParameters.AppServerHost, "192.168.2.51");
    ////    rfc.Add(RfcConfigParameters.Client, "400");
    ////    rfc.Add(RfcConfigParameters.User, "basis");
    ////    rfc.Add(RfcConfigParameters.Password, "support32");
    ////    rfc.Add(RfcConfigParameters.SystemNumber, "00");
    ////    rfc.Add(RfcConfigParameters.Language, "EN");
    ////    //Server Develop
    ////    //rfc.Add(RfcConfigParameters.AppServerHost, "192.168.2.53");
    ////    //rfc.Add(RfcConfigParameters.Client, "220");
    ////    //Server Production
    ////    //rfc.Add(RfcConfigParameters.AppServerHost, "192.168.2.52");
    ////    //rfc.Add(RfcConfigParameters.Client, "300");
    ////    //Server Real
    ////    //rfc.Add(RfcConfigParameters.AppServerHost, "192.168.2.51");
    ////    //rfc.Add(RfcConfigParameters.Client, "400");
    ////    //rfc.Add(RfcConfigParameters.User, "basis");
    ////    //rfc.Add(RfcConfigParameters.Password, "support32");
    ////    //rfc.Add(RfcConfigParameters.SystemNumber, "00");
    ////    //rfc.Add(RfcConfigParameters.Language, "EN");

    ////    RfcDestination rfcDest = RfcDestinationManager.GetDestination(rfc);
    ////    RfcRepository rfcRep = rfcDest.Repository;
    ////    IRfcFunction function = rfcRep.CreateFunction("ZISD0001_POST_SO_DOC");

    ////    function["I_TEST"].SetValue("");

    ////    //1)Zisd0001HdSoShipTo
    ////    DataTable dtShipTo = dsData.Tables["dtShipTo"];
    ////    IRfcStructure tbShipTo = function.GetStructure("IS_ADDR_SHIP_TO");
    ////    //IRfcTable tbShipTo = function.GetTable("IS_ADDR_SHIP_TO");
    ////    //tbShipTo.Append();      
    ////    tbShipTo.SetValue("NAME", dtShipTo.Rows[0]["NAME"]);
    ////    tbShipTo.SetValue("NAME_2", dtShipTo.Rows[0]["NAME_2"]);
    ////    tbShipTo.SetValue("NAME_3", dtShipTo.Rows[0]["NAME_3"]);
    ////    tbShipTo.SetValue("NAME_4", dtShipTo.Rows[0]["NAME_4"]);
    ////    tbShipTo.SetValue("STREET_LNG", dtShipTo.Rows[0]["STREET_LNG"]);
    ////    tbShipTo.SetValue("STR_SUPPL1", dtShipTo.Rows[0]["STR_SUPPL1"]);
    ////    tbShipTo.SetValue("DISTRICT", dtShipTo.Rows[0]["DISTRICT"]);
    ////    tbShipTo.SetValue("CITY", dtShipTo.Rows[0]["CITY"]);
    ////    tbShipTo.SetValue("POSTL_COD1", dtShipTo.Rows[0]["POSTL_COD1"]);
    ////    tbShipTo.SetValue("COUNTRY", dtShipTo.Rows[0]["COUNTRY"]);
    ////    tbShipTo.SetValue("TEL1_NUMBR", dtShipTo.Rows[0]["TEL1_NUMBR"]);
    ////    tbShipTo.SetValue("TEL1_EXT", dtShipTo.Rows[0]["TEL1_EXT"]);
    ////    tbShipTo.SetValue("FAX_NUMBER", dtShipTo.Rows[0]["FAX_NUMBER"]);
    ////    tbShipTo.SetValue("FAX_EXTENS", dtShipTo.Rows[0]["FAX_EXTENS"]);
    ////    tbShipTo.SetValue("E_MAIL", dtShipTo.Rows[0]["E_MAIL"]);
    ////    tbShipTo.SetValue("CONTACT_PERSON", dtShipTo.Rows[0]["CONTACT_PERSON"]);
    ////    tbShipTo.SetValue("TRANSPZONE", dtShipTo.Rows[0]["TRANSPZONE"]);
    ////    function.SetValue("IS_ADDR_SHIP_TO", tbShipTo);

    ////    //2)Zisd0001HdSoSoldTo
    ////    DataTable dtSoldTo = dsData.Tables["dtSoldTo"];
    ////    IRfcStructure tbSoldTo = function.GetStructure("IS_ADDR_SOLD_TO");
    ////    //IRfcTable tbSoldTo = function.GetTable("IS_ADDR_SOLD_TO");
    ////    //tbSoldTo.Append();
    ////    tbSoldTo.SetValue("NAME", dtSoldTo.Rows[0]["NAME"]);
    ////    tbSoldTo.SetValue("NAME_2", dtSoldTo.Rows[0]["NAME_2"]);
    ////    tbSoldTo.SetValue("NAME_3", dtSoldTo.Rows[0]["NAME_3"]);
    ////    tbSoldTo.SetValue("NAME_4", dtSoldTo.Rows[0]["NAME_4"]);
    ////    tbSoldTo.SetValue("STREET_LNG", dtSoldTo.Rows[0]["STREET_LNG"]);
    ////    tbSoldTo.SetValue("STR_SUPPL1", dtSoldTo.Rows[0]["STR_SUPPL1"]);
    ////    tbSoldTo.SetValue("DISTRICT", dtSoldTo.Rows[0]["DISTRICT"]);
    ////    tbSoldTo.SetValue("CITY", dtSoldTo.Rows[0]["CITY"]);
    ////    tbSoldTo.SetValue("POSTL_COD1", dtSoldTo.Rows[0]["POSTL_COD1"]);
    ////    tbSoldTo.SetValue("COUNTRY", dtSoldTo.Rows[0]["COUNTRY"]);
    ////    tbSoldTo.SetValue("TEL1_NUMBR", dtSoldTo.Rows[0]["TEL1_NUMBR"]);
    ////    tbSoldTo.SetValue("TEL1_EXT", dtSoldTo.Rows[0]["TEL1_EXT"]);
    ////    tbSoldTo.SetValue("FAX_NUMBER", dtSoldTo.Rows[0]["FAX_NUMBER"]);
    ////    tbSoldTo.SetValue("FAX_EXTENS", dtSoldTo.Rows[0]["FAX_EXTENS"]);
    ////    tbSoldTo.SetValue("E_MAIL", dtSoldTo.Rows[0]["E_MAIL"]);
    ////    tbSoldTo.SetValue("CONTACT_PERSON", dtSoldTo.Rows[0]["CONTACT_PERSON"]);
    ////    tbSoldTo.SetValue("TAX_ID", dtSoldTo.Rows[0]["TAX_ID"]);
    ////    tbSoldTo.SetValue("BRANCH_NO", dtSoldTo.Rows[0]["BRANCH_NO"]);
    ////    function.SetValue("IS_ADDR_SOLD_TO", tbSoldTo);

    ////    //3)Zisd0001HdFi(ส่วนมัดจำ)
    ////    DataTable dtHdFi = dsData.Tables["dtHdFi"];
    ////    IRfcStructure tbHdFi = function.GetStructure("IS_HD_FI");
    ////    //IRfcTable tbHdFi = function.GetTable("IS_HD_FI");
    ////    //tbHdFi.Append();
    ////    tbHdFi.SetValue("BLDAT", dtHdFi.Rows[0]["BLDAT"]);
    ////    tbHdFi.SetValue("NEWKO", dtHdFi.Rows[0]["NEWKO"]);
    ////    tbHdFi.SetValue("FWBAS", dtHdFi.Rows[0]["FWBAS"]);
    ////    tbHdFi.SetValue("WRBTR", dtHdFi.Rows[0]["WRBTR"]);
    ////    function.SetValue("IS_HD_FI", tbHdFi);

    ////    //4)Zisd0001HdSo
    ////    DataTable dtHdSo = dsData.Tables["dtHdSo"];
    ////    IRfcStructure tbHdSo = function.GetStructure("IS_HD_SO");
    ////    //IRfcTable tbHdSo = function.GetTable("IS_HD_SO");
    ////    //tbHdSo.Append();
    ////    tbHdSo.SetValue("HD_INDICATOR", dtHdSo.Rows[0]["HD_INDICATOR"]);
    ////    tbHdSo.SetValue("PURCH_NO_S", dtHdSo.Rows[0]["PURCH_NO_S"]);
    ////    tbHdSo.SetValue("PO_DAT_S", dtHdSo.Rows[0]["PO_DAT_S"]);
    ////    tbHdSo.SetValue("DOC_TYPE", dtHdSo.Rows[0]["DOC_TYPE"]);
    ////    tbHdSo.SetValue("SALES_ORG", dtHdSo.Rows[0]["SALES_ORG"]);
    ////    tbHdSo.SetValue("DISTR_CHAN", dtHdSo.Rows[0]["DISTR_CHAN"]);
    ////    tbHdSo.SetValue("DIVISION", dtHdSo.Rows[0]["DIVISION"]);
    ////    tbHdSo.SetValue("REQ_DATE_H", dtHdSo.Rows[0]["REQ_DATE_H"]);
    ////    tbHdSo.SetValue("PURCH_NO_C", dtHdSo.Rows[0]["PURCH_NO_C"]);
    ////    tbHdSo.SetValue("PLANT", dtHdSo.Rows[0]["PLANT"]);
    ////    tbHdSo.SetValue("SHIP_COND", dtHdSo.Rows[0]["SHIP_COND"]);
    ////    tbHdSo.SetValue("CURRENCY", dtHdSo.Rows[0]["CURRENCY"]);
    ////    function.SetValue("IS_HD_SO", tbHdSo);

    ////    //5)Zisd0001HdSoPartner
    ////    DataTable dtHdSoPartner = dsData.Tables["dtHdSoPartner"];
    ////    IRfcStructure tbHdSoPartner = function.GetStructure("IS_PARTNER");
    ////    //IRfcTable tbHdSoPartner = function.GetTable("IS_PARTNER");
    ////    //tbHdSoPartner.Append();
    ////    tbHdSoPartner.SetValue("SOLD_TO_ID", dtHdSoPartner.Rows[0]["SOLD_TO_ID"]);
    ////    tbHdSoPartner.SetValue("SHIP_TO_ID", dtHdSoPartner.Rows[0]["SHIP_TO_ID"]);
    ////    tbHdSoPartner.SetValue("EMPLOYEE_ID", dtHdSoPartner.Rows[0]["EMPLOYEE_ID"]);
    ////    function.SetValue("IS_PARTNER", tbHdSoPartner);


    ////    //6)Zisd0001ItSo
    ////    DataTable dtItSo = dsData.Tables["dtItSo"];
    ////    IRfcTable tbItSo = function.GetTable("T_IT_SO");
    ////    for (int i = 0; i < dtItSo.Rows.Count; i++)
    ////    {
    ////        tbItSo.Append();
    ////        tbItSo.SetValue("IT_INDICATOR", dtItSo.Rows[i]["IT_INDICATOR"]);
    ////        tbItSo.SetValue("POITM_NO_S", dtItSo.Rows[i]["POITM_NO_S"]);
    ////        tbItSo.SetValue("MATERIAL", dtItSo.Rows[i]["MATERIAL"]);
    ////        tbItSo.SetValue("ITM_TYPE_USAGE", dtItSo.Rows[i]["ITM_TYPE_USAGE"]);
    ////        tbItSo.SetValue("COND_VALUE", dtItSo.Rows[i]["COND_VALUE"]);
    ////        tbItSo.SetValue("REQ_QTY", dtItSo.Rows[i]["REQ_QTY"]);
    ////    }
    ////    function.SetValue("T_IT_SO", tbItSo);

    ////    //Call SAP Function     
    ////    SODoc_CreateDatatable();
    ////    try
    ////    {
    ////        //execute the function
    ////        function.Invoke(rfcDest);
    ////        dtSOException.Rows.Add("Y");
    ////    }
    ////    catch (Exception e)
    ////    {
    ////        dtSOException.Rows.Add(e.Message);
    ////    }

    ////    IRfcTable tbReturn = function.GetTable("T_Return");

    ////    //----dtSODoc----//
    ////    DataRow drSODoc = dtSODoc.NewRow();
    ////    drSODoc["E_SALES_DOC"] = function.GetValue("E_SALES_DOC").ToString();
    ////    drSODoc["E_ADVANCE_REC_DOC"] = function.GetValue("E_ADVANCE_REC_DOC").ToString();
    ////    drSODoc["E_COMP_CODE"] = function.GetValue("E_COMP_CODE").ToString();
    ////    drSODoc["E_FISCAL_YEAR"] = function.GetValue("E_FISCAL_YEAR").ToString();
    ////    dtSODoc.Rows.Add(drSODoc);

    ////    //Create Dataset
    ////    return SODoc_CreateDataset(tbReturn, dtSODoc, dtSOException);
    ////}
    #endregion


}
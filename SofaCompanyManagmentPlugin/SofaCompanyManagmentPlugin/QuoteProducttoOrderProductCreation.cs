using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Crm.Sdk;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace SofaCompanyManagmentPlugin
{
    public class QuoteProducttoOrderProductCreation : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
            {
                tracingService.Trace("Execution Started");
                try
                {
                    #region declaring Variable

                    #endregion

                    Entity _order = (Entity)context.PostEntityImages["PostImage"];
                   

                    if (_order != null && _order.LogicalName == "salesorder" && _order.Contains("quoteid"))                       
                    {
                       
                        EntityReference _quoteid = _order.GetAttributeValue<EntityReference>("quoteid");

                        #region FetchXML
                        var fetch = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
  <entity name='quotedetail'>
    <attribute name='productid' />
    <attribute name='productdescription' />
    <attribute name='priceperunit' />
    <attribute name='quantity' />
    <attribute name='extendedamount' />
    <attribute name='quotedetailid' />
    <attribute name='volumediscountamount_base' />
    <attribute name='volumediscountamount' />
    <attribute name='uomid' />
    <attribute name='tax_base' />
    <attribute name='tax' />
    <attribute name='shipto_postalcode' />
    <attribute name='shipto_line3' />
    <attribute name='shipto_line2' />
    <attribute name='shipto_line1' />
    <attribute name='shipto_stateorprovince' />
    <attribute name='shipto_telephone' />
    <attribute name='shipto_name' />
    <attribute name='shipto_fax' />
    <attribute name='shipto_country' />
    <attribute name='shipto_contactname' />
    <attribute name='shipto_city' />
    <attribute name='willcall' />
    <attribute name='sequencenumber' />
    <attribute name='isproductoverridden' />
    <attribute name='salesrepid' />
    <attribute name='requestdeliveryby' />
    <attribute name='overriddencreatedon' />
    <attribute name='quoteid' />
    <attribute name='propertyconfigurationstatus' />
    <attribute name='producttypecode' />
    <attribute name='productname' />
    <attribute name='priceperunit_base' />
    <attribute name='ispriceoverridden' />
    <attribute name='parentbundleidref' />
    <attribute name='parentbundleid' />
    <attribute name='quotedetailname' />
    <attribute name='modifiedon' />
    <attribute name='modifiedonbehalfby' />
    <attribute name='modifiedby' />
    <attribute name='manualdiscountamount_base' />
    <attribute name='manualdiscountamount' />
    <attribute name='lineitemnumber' />
    <attribute name='clt_legs_rly' />
    <attribute name='clt_legs' />
    <attribute name='shipto_freighttermscode' />
    <attribute name='clt_fabric_rly' />
    <attribute name='clt_fabric' />
    <attribute name='extendedamount_base' />
    <attribute name='exchangerate' />
    <attribute name='description' />
    <attribute name='transactioncurrencyid' />
    <attribute name='createdon' />
    <attribute name='createdonbehalfby' />
    <attribute name='createdby' />
    <attribute name='productassociationid' />
    <attribute name='baseamount_base' />
    <attribute name='baseamount' />
    <order attribute='productid' descending='false' />
    <filter type='and'>
      <condition attribute='quoteid' operator='eq' uiname='Price Calculation' uitype='quote' value='" + _quoteid.Id + @" ' />
    </filter>
  </entity>
</fetch>";
                        #endregion FetchXML

                        EntityCollection entitycollection = service.RetrieveMultiple(new FetchExpression(fetch));
                        tracingService.Trace("Quote Product Retrieved count : " + entitycollection.Entities.Count.ToString());
                  
                        if (entitycollection.Entities.Count > 0)
                        {
                           
                            foreach (Entity _quoteP in entitycollection.Entities)
                            {
                                tracingService.Trace("Order Product Creation Started");
                                #region Create Nearest Facilities
                                Entity _orderProduct = new Entity("salesorderdetail");
                                _orderProduct.Attributes["isproductoverridden"] = _quoteP.GetAttributeValue<bool>("isproductoverridden");
                                tracingService.Trace("Order Product 1");
                                _orderProduct.Attributes["productdescription"] = _quoteP.GetAttributeValue<string>("productdescription");
                                tracingService.Trace("Order Product 2");
                                _orderProduct.Attributes["productid"] = _quoteP.GetAttributeValue<EntityReference>("productid");
                                tracingService.Trace("Order Product 3");
                                _orderProduct.Attributes["uomid"] = _quoteP.GetAttributeValue<EntityReference>("uomid");
                                tracingService.Trace("Order Product 4");
                                _orderProduct.Attributes["ispriceoverridden"] = _quoteP.GetAttributeValue<bool>("ispriceoverridden");
                                tracingService.Trace("Order Product 5");
                                _orderProduct.Attributes["priceperunit"] = _quoteP.GetAttributeValue<Money>("priceperunit");
                                tracingService.Trace("Order Product 6");
                                _orderProduct.Attributes["volumediscountamount"] = _quoteP.GetAttributeValue<Money>("volumediscountamount");
                                tracingService.Trace("Order Product 7");
                                _orderProduct.Attributes["quantity"] = _quoteP.GetAttributeValue<decimal>("quantity");
                                tracingService.Trace("Order Product 8");
                                _orderProduct.Attributes["baseamount"] = _quoteP.GetAttributeValue<Money>("baseamount");
                                tracingService.Trace("Order Product 9");
                                _orderProduct.Attributes["manualdiscountamount"] = _quoteP.GetAttributeValue<Money>("manualdiscountamount");
                                tracingService.Trace("Order Product 10");
                                _orderProduct.Attributes["tax"] = _quoteP.GetAttributeValue<Money>("tax");
                                tracingService.Trace("Order Product 11");
                                _orderProduct.Attributes["extendedamount"] = _quoteP.GetAttributeValue<Money>("extendedamount");
                                tracingService.Trace("Order Product 12");
                                if (_quoteP.Contains("requestdeliveryby"))
                                {
                                    _orderProduct.Attributes["requestdeliveryby"] = _quoteP.GetAttributeValue<DateTime>("requestdeliveryby");
                                }
                                tracingService.Trace("Order Product 13");
                                _orderProduct.Attributes["salesrepid"] = _quoteP.GetAttributeValue<EntityReference>("salesrepid");
                                tracingService.Trace("Order Product 14");
                                //    _orderProduct.Attributes["quantityshipped"] = _quoteP.GetAttributeValue<decimal>("quantityshipped");
                                // _orderProduct.Attributes["quantitybackordered"] = _quoteP.GetAttributeValue<decimal>("quantitybackordered");
                                // _orderProduct.Attributes["quantitycancelled"] = _quoteP.GetAttributeValue<decimal>("quantitycancelled");
                                _orderProduct.Attributes["willcall"] = _quoteP.GetAttributeValue<bool>("willcall");
                                tracingService.Trace("Order Product 15");
                                _orderProduct.Attributes["shipto_name"] = _quoteP.GetAttributeValue<string>("shipto_name");
                                _orderProduct.Attributes["shipto_line1"] = _quoteP.GetAttributeValue<string>("shipto_line1");
                                _orderProduct.Attributes["shipto_line2"] = _quoteP.GetAttributeValue<string>("shipto_line2");
                                _orderProduct.Attributes["shipto_line3"] = _quoteP.GetAttributeValue<string>("shipto_line3");
                                _orderProduct.Attributes["shipto_city"] = _quoteP.GetAttributeValue<string>("shipto_city");
                                _orderProduct.Attributes["shipto_stateorprovince"] = _quoteP.GetAttributeValue<string>("shipto_stateorprovince");
                                _orderProduct.Attributes["shipto_postalcode"] = _quoteP.GetAttributeValue<string>("shipto_postalcode");
                                _orderProduct.Attributes["shipto_country"] = _quoteP.GetAttributeValue<string>("shipto_country");
                                _orderProduct.Attributes["shipto_telephone"] = _quoteP.GetAttributeValue<string>("shipto_telephone");
                                _orderProduct.Attributes["shipto_fax"] = _quoteP.GetAttributeValue<string>("shipto_fax");
                                tracingService.Trace("Order Product 16");
                                _orderProduct.Attributes["shipto_freighttermscode"] = _quoteP.GetAttributeValue<OptionSetValue>("shipto_freighttermscode");
                                tracingService.Trace("Order Product 17");
                                _orderProduct.Attributes["shipto_contactname"] = _quoteP.GetAttributeValue<string>("shipto_contactname");
                                _orderProduct.Attributes["salesorderid"] = new EntityReference("salesorder",_order.Id);
                                _orderProduct.Attributes["quotedetailid"] = new EntityReference("quotedetail", _quoteP.Id);
                                _orderProduct.Attributes["productname"] = _quoteP.GetAttributeValue<string>("productname");
                                tracingService.Trace("Order Product 18");
                                _orderProduct.Attributes["producttypecode"] = _quoteP.GetAttributeValue<OptionSetValue>("producttypecode");
                                tracingService.Trace("Order Product 19");
                                _orderProduct.Attributes["lineitemnumber"] = _quoteP.GetAttributeValue<int>("lineitemnumber");
                                tracingService.Trace("Order Product 20");
                                _orderProduct.Attributes["transactioncurrencyid"] = _quoteP.GetAttributeValue<EntityReference>("transactioncurrencyid");//description
                                _orderProduct.Attributes["description"] = _quoteP.GetAttributeValue<string>("description");
                                _orderProduct.Attributes["exchangerate"] = _quoteP.GetAttributeValue<decimal>("exchangerate");
                                _orderProduct.Attributes["salesorderdetailname"] = _quoteP.GetAttributeValue<string>("quotedetailname");

                                Guid _nearestFacilityId = service.Create(_orderProduct);
                                tracingService.Trace("Order Product Created");
                                #endregion
                            }
                        }


                    }
                }
                catch (Exception ex)
                {

                    throw new InvalidPluginExecutionException(ex.Message);
                }
            }

        }

    }
}

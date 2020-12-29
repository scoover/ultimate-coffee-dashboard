function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Ultimate Coffee')
    .addItem('Update from Shopify', 'updateFromShopify')
    .addToUi();
};

function updateFromShopify() {
  updateCustomers();
  updateOrders();
}

function updateCustomers() {
  // load the linked spreadsheet object
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var customersSheet = spreadsheet.getSheetByName('Customers');

  var results =  shopifyApiPost(
    '/admin/api/2020-10/customers.json?fields=id,created_at,updated_at,first_name,last_name,total_spent,orders_count,last_order_name,addresses', null);
  var rows = [ ];

  if(results.hasOwnProperty('customers') && results.customers.length) {
    results.customers.forEach(function (customer, index) {
      var customerInfo = [ ];

      // skip customers that have zero orders
      if(!customer.hasOwnProperty('orders_count') || !customer.orders_count) { return; }

      customerInfo.push(getProperty(customer,['id']));
      customerInfo.push(getProperty(customer,['first_name'])? getProperty(customer,['first_name'])[0]:null);
      customerInfo.push(getProperty(customer,['last_name']));
      customerInfo.push(new Date(getProperty(customer,['created_at'])));
      customerInfo.push(new Date(getProperty(customer,['updated_at'])));
      customerInfo.push(getProperty(customer,['total_spent']));
      customerInfo.push(getProperty(customer,['orders_count']));
      customerInfo.push(getProperty(customer,['last_order_name']));
      
      var zips = '';
      if(customer.hasOwnProperty('addresses') && customer.addresses.length) {
        customer.addresses.forEach(function (a) {
          if(a.hasOwnProperty('zip') && a.zip) {
            zips += (zips.length? ' ':'') + a.zip
          }
        });
      }
      
      customerInfo.push('\''+zips);
      rows.push(customerInfo);
    });

    customersSheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  }
}

function updateOrders() {
  // load the linked spreadsheet object
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var itemsSheet = spreadsheet.getSheetByName('Order Items');

  var results =  shopifyApiPost(
    '/admin/api/2020-10/orders.json?status=any&financial_status=paid&fields=id,order_number,line_items', null);
  var rows = [ ];

  if(results.hasOwnProperty('orders') && results.orders.length) {
    results.orders.forEach(function (o, index) {

      // skip orders that have zero line items
      if(!o.hasOwnProperty('line_items') || !o.line_items.length) { return; }

      o.line_items.forEach(function (i) {
        var itemInfo = [ ];
        itemInfo.push(getProperty(o,['id']));
        itemInfo.push(getProperty(o,['order_number']));
        itemInfo.push(getProperty(i,['quantity']));
        itemInfo.push(getProperty(i,['vendor']));
        itemInfo.push(getProperty(i,['title']));      
        rows.push(itemInfo);  
      });
    });

    itemsSheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  }
}

function shopifyApiPost(apiCall, postData) {
  var request = 'https://ultimate-coffee-program.myshopify.com' + apiCall;
  var options = {
    headers: { "Authorization": "Basic " + Utilities.base64Encode(PropertiesService.getScriptProperties().getProperty('SHOPIFY')) },
    contentType: 'application/json',
  };

  if(postData) {
    options.method = 'post';
    options.payload = JSON.stringify(postData);
  }
  else {
    options.method = 'get';
  }

  var response = UrlFetchApp.fetch(request, options);

  if (response.getResponseCode() === 200) {
    try { return JSON.parse(response.getContentText()); }
    catch(e) { }
  }

  return null;
}

// getProperty(object, property array): get value from array of nested property names (null if not present)
function getProperty(o, p) {
  var c = JSON.parse(JSON.stringify(o));
  for(var i = 0; i < p.length; i++) {
    if(!c.hasOwnProperty(p[i])) {
      return null;
    }
    c = c[p[i]];
  }
  return c;
}
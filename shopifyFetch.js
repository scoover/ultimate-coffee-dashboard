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

  var results =  shopifyApiPostRequest(
    '/admin/api/2020-10/customers.json',
    {
      fields: 'id,created_at,updated_at,first_name,last_name,total_spent,orders_count,last_order_name,addresses',
      limit: 250
    },
    'customers'
  );

  var rows = [ ];

  if(results.length) {
    results.forEach(function (customer, index) {
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

  var results =  shopifyApiPostRequest(
    '/admin/api/2020-10/orders.json',
    {
      status: 'any',
      financial_status: 'paid',
      fields: 'id,order_number,line_items',
      limit: 250
    },
    'orders'
  );
  
  var rows = [ ];

  if(results.length) {
    results.forEach(function (o, index) {

      // skip orders that have zero line items
      if(!o.hasOwnProperty('line_items') || !o.line_items.length) { return; }

      o.line_items.forEach(function (i) {
        var discount = 0;

        if(i.hasOwnProperty('discount_allocations')) {
          i.discount_allocations.forEach(function (d) {
            if(d.hasOwnProperty('amount')) { discount += d.amount; }
          });
        };

        var itemInfo = [ ];
        itemInfo.push(getProperty(o,['id']));
        itemInfo.push(getProperty(o,['order_number']));
        itemInfo.push(getProperty(i,['quantity']));

        var vendor = getProperty(i,['vendor']);

        // clean up the vendor name if it is null
        if(!vendor) {
          var lookup = {
            '35137261633699': 'Olympia Coffee',
            '35125657993379': 'Olympia Coffee',
            '35125709602979': 'Lighthouse Roasters'
          }

          var variantId = getProperty(i,['variant_id']);

          if(variantId && lookup.hasOwnProperty(variantId)) {
            vendor = lookup[variantId];
          }          
        }

        itemInfo.push(vendor);
        itemInfo.push(getProperty(i,['title']));      
        itemInfo.push(getProperty(i,['price']));
        itemInfo.push(discount);
        rows.push(itemInfo);  
      });
    });

    itemsSheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  }
}

function shopifyApiPostRequest(endpoint, params, responseKey) {
  var results = [ ];
  var response = null;

  do {
    var apiCall = endpoint;

    if((response !== null) && response.hasOwnProperty('next')) {
      // add the pagination field `page_info`
      // https://shopify.dev/tutorials/make-paginated-requests-to-rest-admin-api
      params.page_info = response.next;
      
      // remove parameters that are invalid after initial request
      Object.keys(params).forEach(function (param) {      
        if(['page_info', 'limit', 'fields'].indexOf(param)) { delete params[param]; }
      });
    }

    // build the api call by adding parameters
    Object.keys(params).forEach(function (param, index) {      
      apiCall += (index? '&':'?') + param + '=' + params[param];
    });

    // make the Post request
    response = shopifyApiPost(apiCall, null);

    // append the content if present
    if(response.hasOwnProperty('content') && response.content.hasOwnProperty(responseKey)) {
      results = results.concat(response.content[responseKey]);
    }
  } while (response.hasOwnProperty('next'));

  return results;
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
    try { 
      var output = { content: JSON.parse(response.getContentText()) };
      var paging = shopifyApiGetPaging(response);
      if(paging.hasOwnProperty('next')) { output.next = paging.next; }
      if(paging.hasOwnProperty('previous')) { output.previous = paging.previous; }
      return output;
    }
    catch(e) { }
  }

  return { content: null };
}

function shopifyApiGetPaging(response) {
  var paging = { };
  if (response.getAllHeaders().Link) {
    var match;
    while ((match = /page_info=([^>]+)>; rel="([^"]+)/g.exec(response.getAllHeaders().Link)) !== null) {
      // match[1] is the cursor (page_info)
      // match[2] is "next" or "previous"
      paging[match[2]] = match[1];
    }
  }
  // returns {next: 'the next cursor', previous: 'the previous cursor'}
  return paging;
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
// Create a private App in Shopify admin and insert app password below
// Sheets need to be created before running with columns of desired output set
// Use keys found in Shopify Rest API Order Object as row headers
// https://shopify.dev/api/admin-rest/2022-04/resources/order#resource-object

// Supports the following sheets/order objects with default columns -
// orders - id, number, order_name, name, current_total_discounts, current_total_price, current_subtotal_price, current_total_tax, financial_status, landing_site, created_at, processed_at, processing_method, subtotal_price, source_url, tags, token, total_discounts, total_line_items_price, total_price, total_tax, updated_at
// shipping_lines - order_id, ordered_at, code, price, title, source, discounted_price, carrier_identifier
// discount_applications - order_id, created_at, type, title, value, value_type, code, description, target_type, target_selection, allocation_method, ordered_at
// refunds = id, order_id, ordered_at, created_at, note, transaction_id, kind, amount
// line_items = id, order_id, sku, name, price, title, vendor, quantity, product_id, variant_id, variant_title, total_discount, fulfillment_status, ordered_at
// customer - order_id, ordered_at, id, email, phone, last_name, created_at, first_name, admin_graphql_api_id, updated_at, orders_count, last_order_id

const password = YOUR_SHOPIFY_PRIVATE_APP_PASSWORD;
const ss = SpreadSheetApp.getActiveSpreadsheet();

const shop_base = 'https://YOUR-STORE.myshopify.com'

const endpoint = '/admin/api/2022-04/orders.json'

const today = new Date();
const thirty_days = new Date(today.getFullYear(), today.getMonth()-1, today.getDate());
const start_str = Utilities.formatDate(thirty_days, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "yyyy-MM-dd");
const end_str = Utilities.formatDate(today, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "yyyy-MM-dd")


// Can accept start and end date in "yyyy-MM-dd" format or it will default to pull all data up to 30 days ago
function RunOrders(start='', end='') {

    if (start === '' || end === '') {
        start = start_str;
        end = end_str;
    }
    const date_exp = /^\d{4}[\-](0?[1-9]|1[012])[\-](0?[1-9]|[12][0-9]|3[01])$/;
    if (!date_exp.test(start) || !date_exp.test(end)) {
        Logger.log("Dates must be in yyyy-MM-dd format");
        return
    }
    GetOrders(start, end);

}



// Main Function
function GetOrders(start, end) {

    var params = {
        created_at_min: start,
        created_at_max: end,
        fields: getColumnList()
    }

    var done = false;
    var next_link = ''
    do  {
        if (next_link === '') {
            var [resp, headers] = shopifyAPI(endpoint, params);
        } else {
            var [resp, headers] = shopifyAPI(endpoint, params, next_link)
        }
        var orders_arr = resp['orders']
        var cols = getColumnList()
        var numColumns = cols.length;
        var respKeys = Object.keys(orders_arr[0]);
        var respCols = respKeys.length;
        if (cols.filter(x => respKeys.indexOf(x) === -1).length !== 0) {
            Logger.log("Keys missing from return - " + cols.filter(x => respKeys.indexOf(x) === -1));
            return
        }
        ordersTable(orders_arr);
        refundsTable(orders_arr);
        lineItemsTable(orders_arr);
        customerTable(orders_arr);
        discountsTable(orders_arr);
        shippingTable(orders_arr);

        if (headers["Link"] !== undefined) {
            var links = parseLinkHeader(headers['Link']);
            Logger.log(headers['Link']);
            if (links['next'] === undefined) {
                done = true;
            } else {
                next_link = links["next"]['href']
                Logger.log(next_link)
            }
        } else {
            done = true;
        }

    } while (!done);

}

function customerTable(orders_arr) {
    var rng_arr = makeCustomerArray(orders_arr);
    removeDuplicates(rng_arr, customer, "order_id");
    SpreadsheetApp.flush();
    appendRows(customer, rng_arr);
    SpreadsheetApp.flush();
    processDataRange(customer);
}

function discountsTable(orders_arr) {
    var rng_arr = makeDiscountsArray(orders_arr);
    if (rng_arr.length > 0){
        removeAllDuplicates(rng_arr, discount_applications, "order_id")
        SpreadsheetApp.flush();
        appendRows(discount_applications, rng_arr);
        SpreadsheetApp.flush();
        processDataRange(discount_applications);}
}

function ordersTable(orders_arr) {
    var rng_arr = makeArray(getOrderHeaders(), orders_arr);
    removeDuplicates(rng_arr, orders, "id");
    SpreadsheetApp.flush();
    appendRows(orders, rng_arr);
    SpreadsheetApp.flush();
    processDataRange(orders);
}

function refundsTable(orders_arr) {
    var rng_arr = makeRefundsArray(orders_arr);
    if (rng_arr.length > 0){
        removeDuplicates(rng_arr, refunds, 'id');
        SpreadsheetApp.flush();
        appendRows(refunds, rng_arr);
        SpreadsheetApp.flush();
        processDataRange(refunds)}
}

function lineItemsTable(orders_arr) {
    var rng_arr = makeLIArray(orders_arr);
    removeDuplicates(rng_arr, line_items, "id");
    SpreadsheetApp.flush();
    appendRows(line_items, rng_arr);
    SpreadsheetApp.flush();
    processDataRange(line_items)
}

function shippingTable(orders_arr) {
    var rng_arr = makeShippingArray(orders_arr);
    if (rng_arr.length > 0){
        removeDuplicates(rng_arr, shipping_lines, 'order_id');
        SpreadsheetApp.flush();
        appendRows(shipping_lines, rng_arr);
        SpreadsheetApp.flush();
        processDataRange(shipping_lines)}
}

function shopifyAPI(endpoint, query, next='') {
    var headers = {
        'X-Shopify-Access-Token': password,
    }
    var options = {
        method: 'get',
        headers: headers
    }
    if (next == ''){
        var url = buildURL(endpoint, query);
        var resp = UrlFetchApp.fetch(url, options);
    } else {
        var url = next;
        var resp = UrlFetchApp.fetch(url, options);
    }

    return [JSON.parse(resp.getContentText()), resp.getHeaders()]
}

function buildURL(endpoint, params = {}) {
    var url = shop_base + endpoint;
    if (params !== undefined && Object.keys(params).length > 0) {
        let i = 0
        for (let p in params) {
            if (i === 0) {
                url += '?' + p + '=' + encodeURIComponent(params[p]);

            } else {
                url += '&' + p + '=' + encodeURIComponent(params[p]);
            }
            i++
        }
    }
    return url
}

function unquote(value) {
    if (value.charAt(0) == '"' && value.charAt(value.length - 1) == '"') return value.substring(1, value.length - 1);
    return value;
}

function parseLinkHeader(header) {
    var linkexp = /<[^>]*>\s*(\s*;\s*[^\(\)<>@,;:"\/\[\]\?={} \t]+=(([^\(\)<>@,;:"\/\[\]\?={} \t]+)|("[^"]*")))*(,|$)/g;
    var paramexp = /[^\(\)<>@,;:"\/\[\]\?={} \t]+=(([^\(\)<>@,;:"\/\[\]\?={} \t]+)|("[^"]*"))/g;

    var matches = header.match(linkexp);
    var rels = {};
    for (let i = 0; i < matches.length; i++) {
        var split = matches[i].split('>');
        var href = split[0].substring(1);
        var ps = split[1];
        var link = {};
        link.href = href;
        var s = ps.match(paramexp);
        for (let j = 0; j < s.length; j++) {
            var p = s[j];
            var paramsplit = p.split('=');
            var name = paramsplit[0];
            link[name] = unquote(paramsplit[1]);
        }

        if (link.rel !== undefined) {
            rels[link.rel] = link;
        }
    }

    return rels;
}
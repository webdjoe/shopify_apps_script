const orders = 'orders';
const line_items = 'line_items';
const customer = 'customer';
const refunds = 'refunds';
const shipping_lines = 'shipping_lines';
const discount_applications = 'discount_applications';

function makeArray(headers, obj_list) {
    var rng_arr = [];
    for (let obj of obj_list) {
        var line_arr = []
        headers.forEach((item) => {
            var line_item = (item in obj) ? obj[item] : '';
            if (Array.isArray(line_item)) {
                line_item = line_item.toString();
            }
            line_arr.push(line_item);
        })
        rng_arr.push(line_arr);
    }
    return rng_arr
}

function makeDiscountsArray(orders_arr) {
    var disc_arr = [];
    var headers = getDiscountHeaders();
    for (let ord of orders_arr) {
        var disc_list = ord.discount_applications;
        if (disc_list.length < 1) {
            continue
        }
        disc_list.forEach( (disc) => {
            var disc_line = [];
            disc['order_id'] = ord['id'];
            disc['ordered_at'] = ord['created_at'];
            headers.forEach( (head) => {
                var itm = (head in disc) ? disc[head] : '';
                disc_line.push(itm);
            })
            disc_arr.push(disc_line);
        })

    }
    return disc_arr
}

function makeCustomerArray(orders_arr) {
    var cust_arr = [];
    var headers = getCustomerHeaders();
    for (let ord of orders_arr) {
        var cust_obj = ord.customer;
        if (typeof cust_obj == 'undefined') {
            continue
        }
        var customer_line = []
        cust_obj['order_id'] = ord['id'];
        cust_obj['ordered_at'] = ord['created_at']
        headers.forEach((head) => {
            var itm = (head in cust_obj) ? cust_obj[head]: '';
            customer_line.push(itm);

        })
        cust_arr.push(customer_line)
    }
    return cust_arr
}

function makeLIArray(orders_arr) {
    var line_arr = [];
    var headers = getLineItemHeaders();
    for (let ord of orders_arr) {
        var line_items_list = ord.line_items;
        if (line_items_list.length < 1) {
            continue
        }
        line_items_list.forEach( (li_obj) => {
            var li_line = []
            li_obj['ordered_at'] = ord['created_at'];
            li_obj['order_id'] = ord['id'];
            headers.forEach((head) => {
                var itm = (head in li_obj) ? li_obj[head] : '';
                li_line.push(itm);
            })
            line_arr.push(li_line);
        })
    }
    return line_arr;
}

function makeShippingArray(orders_arr) {
    var ship_arr = []
    var header = getShippingHeaders();
    for (let ord of orders_arr) {
        var ship_list = ord.shipping_lines;
        if (ship_list.length < 1) {
            continue
        }

        ship_list.forEach( (ship_obj) => {
            var line = []


            ship_obj['ordered_at'] = ord['created_at'];
            ship_obj['order_id'] = ord['id'];

            header.forEach((a) => {
                var line_itm = (a in ship_obj) ? ship_obj[a] : '';
                line.push(line_itm);
            })
            ship_arr.push(line);
        })}
    return ship_arr
}

function makeRefundsArray(orders_arr) {
    var refunds_arr = []
    var header = getRefundHeaders();
    for (let ord of orders_arr) {
        var ref_list = ord.refunds;
        if (ref_list.length < 1) {
            continue
        }

        ref_list.forEach( (ref_obj) => {
            var line = []


            ref_obj['ordered_at'] = ord['created_at'];
            if (ref_obj['transactions'].length == 1) {
                var trans = ref_obj['transactions'][0];
                ref_obj['transaction_id'] = trans['id'];
                ref_obj['amount'] = trans['amount'];
                ref_obj['kind'] = trans['kind'];
            } else { return }
            header.forEach((a) => {
                var line_itm = (a in ref_obj) ? ref_obj[a] : '';
                line.push(line_itm);
            })
            refunds_arr.push(line);
        })}
    return refunds_arr
}

function appendRows(sht_name, arr) {
    var sht = ss.getSheetByName(sht_name);
    last_row = sht.getLastRow();
    sht.getRange(last_row+1, 1, arr.length, arr[0].length).setValues(arr);

}

function processDataRange(sht_name){

    var sh = ss.getSheetByName(sht_name);
    var data_rng = sh.getDataRange().offset(1, 0, sh.getLastRow() - 1);
    var data = data_rng.getValues();

    var targetData = new Array();
    for(n=0;n<data.length;++n){
        if(data[n].join().replace(/,/g,'')!=''){
            targetData.push(data[n])};
    }
    data_rng.clear();
    targetData.sort((a,b) => a[1] - b[1])
    sh.getRange(2,1,targetData.length,targetData[0].length).setValues(targetData);
}

function removeDuplicates(obj_list, sht_name, id = "id") {
    var sht = ss.getSheetByName(sht_name);
    var sht_arr = sht.getDataRange().getValues();
    for (let a = 0; a < sht_arr[0].length; a++) {
        if (sht_arr[0][a] == id) {
            var id_col = a;
            break
        }
    }
    for (let x = 0; x < sht_arr.length; x++) {
        var row_id = sht_arr[x][id_col];
        for (let y = 0; y < obj_list.length; y++) {
            if (obj_list[y][id_col] == row_id) {
                sht.getRange(x+1, 1, 1, sht_arr[0].length).clearContent();
                break
            }
        }
    }

}

function removeAllDuplicates(obj_list, sht_name, id = "id") {
    var sht = ss.getSheetByName(sht_name);
    var sht_arr = sht.getDataRange().getValues();
    for (let a = 0; a < sht_arr[0].length; a++) {
        if (sht_arr[0][a] == id) {
            var id_col = a;
            break
        }
    }
    for (let x = 0; x < sht_arr.length; x++) {
        var row_id = sht_arr[x][id_col];
        for (let y = 0; y < obj_list.length; y++) {
            if (obj_list[y][id_col] == row_id) {
                sht.getRange(x+1, 1, 1, sht_arr[0].length).clearContent();
            }
        }
    }

}

function getColumnList() {
    var order_list = getOrderHeaders();
    order_list.push(line_items, customer, refunds, discount_applications, shipping_lines);
    return order_list;
}

function getColumns(data) {
    var sht = ss.getSheetByName(data);
    var last_col = sht.getLastColumn();
    var header_list = sht.getRange(1,1,1,last_col).getValues();
    return header_list[0];
}

function getShippingHeaders() {
    return getColumns(shipping_lines)
}

function getCustomerHeaders() {
    return getColumns(customer)
}

function getLineItemHeaders() {
    return getColumns(line_items)
}

function getOrderHeaders() {
    return getColumns(orders)
}

function getDiscountHeaders() {
    return getColumns(discount_applications)
}

function getRefundHeaders() {
    return getColumns(refunds)
}
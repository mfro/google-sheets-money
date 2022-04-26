'use strict';

var exports = {};

/*! *****************************************************************************
Copyright (c) Microsoft Corporation.

Permission to use, copy, modify, and/or distribute this software for any
purpose with or without fee is hereby granted.

THE SOFTWARE IS PROVIDED "AS IS" AND THE AUTHOR DISCLAIMS ALL WARRANTIES WITH
REGARD TO THIS SOFTWARE INCLUDING ALL IMPLIED WARRANTIES OF MERCHANTABILITY
AND FITNESS. IN NO EVENT SHALL THE AUTHOR BE LIABLE FOR ANY SPECIAL, DIRECT,
INDIRECT, OR CONSEQUENTIAL DAMAGES OR ANY DAMAGES WHATSOEVER RESULTING FROM
LOSS OF USE, DATA OR PROFITS, WHETHER IN AN ACTION OF CONTRACT, NEGLIGENCE OR
OTHER TORTIOUS ACTION, ARISING OUT OF OR IN CONNECTION WITH THE USE OR
PERFORMANCE OF THIS SOFTWARE.
***************************************************************************** */

var __assign = function() {
    __assign = Object.assign || function __assign(t) {
        for (var s, i = 1, n = arguments.length; i < n; i++) {
            s = arguments[i];
            for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p)) t[p] = s[p];
        }
        return t;
    };
    return __assign.apply(this, arguments);
};

function onOpen() {
    SpreadsheetApp.getUi()
        .createMenu('money')
        .addItem('import', import_data.name)
        .addItem('display', display_data.name)
        .addToUi();
}
var Date;
(function (Date) {
    function parse(value) {
        var parts = value.split('/');
        assert(parts.length == 3, "invalid date: " + value);
        var month = parseInt(parts[0]);
        assert(month >= 1 && month <= 12, "invalid date: " + value);
        var day = parseInt(parts[1]);
        assert(day >= 1 && day <= 31, "invalid date: " + value);
        var year = parseInt(parts[2]);
        assert(year >= 1000 && year <= 9999, "invalid date: " + value);
        return { day: day, month: month, year: year };
    }
    Date.parse = parse;
    function unparse(value) {
        var month = value.month.toString();
        var day = value.day.toString();
        var year = value.year.toString();
        month = '0'.repeat(2 - month.length) + month;
        day = '0'.repeat(2 - day.length) + day;
        return month + "/" + day + "/" + year;
    }
    Date.unparse = unparse;
    function eq(a, b) {
        return a.day == b.day
            && a.month == b.month
            && a.year == b.year;
    }
    Date.eq = eq;
})(Date || (Date = {}));
var Money;
(function (Money) {
    function parse(value) {
        if (value == '')
            return { cents: 0 };
        var match = /^\$([\d,]+)\.(\d{2})$/.exec(value);
        assert(match != null, "invalid money: " + value);
        var dollars = parseInt(match[1].replace(/,/g, ''));
        var cents = parseInt(match[2]);
        return { cents: dollars * 100 + cents };
    }
    Money.parse = parse;
    function unparse(value) {
        if (value.cents == 0)
            return '';
        var cents = (value.cents % 100).toString();
        var dollars = Math.floor(value.cents / 100).toString();
        cents = '0'.repeat(2 - cents.length) + cents;
        for (var i = dollars.length - 3; i > 0; i -= 4) {
            dollars = dollars.substr(0, i) + ',' + dollars.substr(i);
        }
        return "$" + dollars + "." + cents;
    }
    Money.unparse = unparse;
    function eq(a, b) {
        return a.cents == b.cents;
    }
    Money.eq = eq;
})(Money || (Money = {}));
var Transaction;
(function (Transaction) {
    function load(sheet, index, count) {
        var data = sheet.getSheetValues(index + 2, 1, count !== null && count !== void 0 ? count : 1, 6);
        var loaded = data.map(parse);
        if (count === undefined)
            return loaded[0];
        return loaded;
    }
    Transaction.load = load;
    function save(sheet, index, transactions) {
        if (!Array.isArray(transactions))
            transactions = [transactions];
        var values = transactions.map(unparse);
        sheet.getRange(index + 2, 1, transactions.length, 6).setValues(values);
    }
    Transaction.save = save;
    function parse(row) {
        var date = Date.parse(row[0]);
        var description = row[1];
        var withdrawn = Money.parse(row[2]);
        var deposited = Money.parse(row[3]);
        var category = row[4];
        var balance = Money.parse(row[5]);
        return {
            date: date,
            description: description,
            withdrawn: withdrawn,
            deposited: deposited,
            category: category,
            balance: balance,
        };
    }
    Transaction.parse = parse;
    function unparse(t) {
        return [
            Date.unparse(t.date),
            t.description,
            Money.unparse(t.withdrawn),
            Money.unparse(t.deposited),
            t.category,
            Money.unparse(t.balance),
        ];
    }
    Transaction.unparse = unparse;
    function eq(a, b) {
        return Date.eq(a.date, b.date)
            && a.description == b.description
            && Money.eq(a.withdrawn, b.withdrawn)
            && Money.eq(a.deposited, b.deposited)
            && Money.eq(a.balance, b.balance);
    }
    Transaction.eq = eq;
})(Transaction || (Transaction = {}));
function import_data() {
    var spread = SpreadsheetApp.getActiveSpreadsheet();
    var import_sheet = spread.getSheetByName('import');
    assert(import_sheet != null, 'no sheet import');
    var raw_sheet = spread.getSheetByName('raw');
    assert(raw_sheet != null, 'no sheet raw');
    var count = import_sheet.getLastRow() - 1; // 1 row of header, the rest is data
    var data = Transaction.load(import_sheet, 0, count);
    var recent = Transaction.load(raw_sheet, 0);
    var recent_index = data.findIndex(function (t) { return Transaction.eq(t, recent); });
    console.log(recent);
    for (var _i = 0, data_1 = data; _i < data_1.length; _i++) {
        var x = data_1[_i];
        console.log(x);
    }
    assert(recent_index != -1, 'no recent :(');
    if (recent_index > 0) {
        data = data.slice(0, recent_index);
        raw_sheet.insertRowsBefore(2, data.length);
        Transaction.save(raw_sheet, 0, data);
        var extra = data.map(function (v, i) { return [
            "=D" + (2 + i) + " - C" + (2 + i),
            "",
            "=TRANSACTION_DESCRIPTION(B" + (2 + i) + ")",
        ]; });
        raw_sheet.getRange(2, 7, data.length, 3).setValues(extra);
        Logger.log("import successful: " + data.length);
        import_sheet.getRange(2, 1, data.length).setBackground('#d9ead3');
    }
}
function display_data() {
    var _a;
    var spread = SpreadsheetApp.getActiveSpreadsheet();
    var raw_sheet = spread.getSheetByName('raw');
    assert(raw_sheet != null, 'no sheet raw');
    var display_sheet = spread.getSheetByName('display');
    assert(display_sheet != null, 'no sheet display');
    var categories_sheet = spread.getSheetByName('categories');
    assert(categories_sheet != null, 'no sheet categories');
    var total = raw_sheet.getLastRow() - 1;
    var data;
    if (display_sheet.getLastRow() <= 2) {
        data = Transaction.load(raw_sheet, 0, total);
    }
    else {
        var rows = display_sheet.getLastRow() - 1;
        var refs = display_sheet.getRange(2, 1, rows, 1)
            .getFormulas()
            .map(function (row) { return row[0]; });
        var previous = refs.find(function (i) { return i !== ''; });
        assert(previous !== undefined, 'invalid state');
        var previous_ref = (_a = /^=raw!A(\d+)$/.exec(previous)) === null || _a === void 0 ? void 0 : _a[1];
        assert(previous_ref !== undefined, 'invalid state');
        var previous_index = parseInt(previous_ref) - 2;
        data = Transaction.load(raw_sheet, 0, previous_index + 1);
        display_sheet.deleteRows(2, 2);
    }
    data.reverse();
    while (data.length > 0) {
        var index = data.findIndex(function (d) { return d.date.month != data[0].date.month || d.date.year != data[0].date.year; });
        if (index == -1)
            index = data.length;
        var group = data.splice(0, index);
        var values = group.map(function (t, i) { return [
            "=raw!A" + (2 + data.length + i),
            "=raw!G" + (2 + data.length + i),
            "=raw!F" + (2 + data.length + i),
            "=raw!H" + (2 + data.length + i),
            "=raw!I" + (2 + data.length + i),
        ]; });
        display_sheet.insertRowsBefore(2, values.length);
        var depth = display_sheet.getRowGroupDepth(2);
        var body = display_sheet.getRange(2, 1, values.length, 5);
        body.setFontWeight(null);
        body.setValues(values);
        body.shiftRowGroupDepth(1 - depth);
        body.setHorizontalAlignments(values.map(function () { return [
            'left',
            'right',
            'right',
            'left',
            'left',
        ]; }));
        display_sheet.getRowGroup(2, 1).collapse();
        var date = __assign(__assign({}, group[0].date), { day: 1 });
        var net = "=SUM(B3:B" + (2 + group.length) + ")";
        var balance = "=C3";
        display_sheet.insertRowBefore(2);
        var header = display_sheet.getRange(2, 1, 1, 4);
        header.setFontWeight('bold');
        header.shiftRowGroupDepth(-1);
        header.setValues([[
                Date.unparse(date),
                net,
                balance,
                ''
            ]]);
        display_sheet.getRange(2, 1).setNumberFormat('mmmm yyy');
    }
}
function TRANSACTION_DESCRIPTION(raw) {
    // let matches: { [key: string]: string[] } = {
    //     'amazon': ['amazon', 'amzn'],
    //     'kroger': ['kroger'],
    //     'michigan pay': ['univ of michigan dir dep'],
    //     'spotify': ['spotify'],
    //     'utowers': ['university tower web pay'],
    // };
    // let lower = raw.toLowerCase();
    // for (let label in matches) {
    //     for (let match of matches[label]) {
    //         if (lower.indexOf(match) != -1) {
    //             return label;
    //         }
    //     }
    // }
    // if (lower.startsWith('debit card purchase')) {
    //     let rest = lower.slice(20);
    //     let parts = rest.split(' ').filter(a => a)
    //     let details = parts[1];
    //     let location = parts[parts.length - 2];
    //     let state = parts[parts.length - 1];
    //     let description = parts.slice(1, parts.length - 2).join(' ');
    //     return `debit: ${description} (${location} ${state})`;
    // }
    // if (lower.startsWith('ach credit')) {
    //     let rest = lower.slice(11);
    //     return `ach credit: ${rest}`;
    // }
    return raw;
}
function assert(condition, message) {
    if (!condition) {
        Logger.log(message);
        throw new Error(message);
    }
}

exports.TRANSACTION_DESCRIPTION = TRANSACTION_DESCRIPTION;
exports.display_data = display_data;
exports.import_data = import_data;
exports.onOpen = onOpen;

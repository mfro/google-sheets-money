type Sheet = GoogleAppsScript.Spreadsheet.Sheet;

export function onOpen() {
    SpreadsheetApp.getUi()
        .createMenu('money')
        .addItem('import', import_data.name)
        .addItem('display', display_data.name)
        .addToUi();
}

interface Date {
    year: number,
    month: number,
    day: number,
}

namespace Date {
    export function parse(value: string): Date {
        let parts = value.split('/');
        assert(parts.length == 3, `invalid date: ${value}`);
        let month = parseInt(parts[0]);
        assert(month >= 1 && month <= 12, `invalid date: ${value}`);
        let day = parseInt(parts[1]);
        assert(day >= 1 && day <= 31, `invalid date: ${value}`);
        let year = parseInt(parts[2]);
        assert(year >= 1000 && year <= 9999, `invalid date: ${value}`);
        return { day, month, year };
    }

    export function unparse(value: Date): string {
        let month = value.month.toString();
        let day = value.day.toString();
        let year = value.year.toString();

        month = '0'.repeat(2 - month.length) + month;
        day = '0'.repeat(2 - day.length) + day;

        return `${month}/${day}/${year}`;
    }

    export function eq(a: Date, b: Date): boolean {
        return a.day == b.day
            && a.month == b.month
            && a.year == b.year;
    }
}

interface Money {
    cents: number,
}

namespace Money {
    export function parse(value: string): Money {
        if (value == '') return { cents: 0 };

        let match = /^\$([\d,]+)\.(\d{2})$/.exec(value);
        assert(match != null, `invalid money: ${value}`);

        let dollars = parseInt(match[1].replace(/,/g, ''));
        let cents = parseInt(match[2]);
        return { cents: dollars * 100 + cents };
    }

    export function unparse(value: Money): string {
        if (value.cents == 0) return '';

        let cents = (value.cents % 100).toString();
        let dollars = Math.floor(value.cents / 100).toString();

        cents = '0'.repeat(2 - cents.length) + cents;
        for (let i = dollars.length - 3; i > 0; i -= 4) {
            dollars = dollars.substr(0, i) + ',' + dollars.substr(i);
        }

        return `$${dollars}.${cents}`;
    }

    export function eq(a: Money, b: Money): boolean {
        return a.cents == b.cents;
    }
}

interface Transaction {
    date: Date,
    description: string,
    withdrawn: Money,
    deposited: Money,
    category: string,
    balance: Money
}

namespace Transaction {
    export function load(sheet: Sheet, index: number): Transaction;
    export function load(sheet: Sheet, index: number, count: number): Transaction[];
    export function load(sheet: Sheet, index: number, count?: number): Transaction | Transaction[] {
        let data = sheet.getSheetValues(index + 2, 1, count ?? 1, 6);
        let loaded = data.map(parse);

        if (count === undefined)
            return loaded[0];
        return loaded;
    }

    export function save(sheet: Sheet, index: number, transactions: Transaction | Transaction[]) {
        if (!Array.isArray(transactions))
            transactions = [transactions];
        let values = transactions.map(unparse);

        sheet.getRange(index + 2, 1, transactions.length, 6).setValues(values);
    }

    export function parse(row: any[]): Transaction {
        let date = Date.parse(row[0]);
        let description = row[1];
        let withdrawn = Money.parse(row[2]);
        let deposited = Money.parse(row[3]);
        let category = row[4];
        let balance = Money.parse(row[5]);

        return {
            date,
            description,
            withdrawn,
            deposited,
            category,
            balance,
        }
    }

    export function unparse(t: Transaction): any[] {
        return [
            Date.unparse(t.date),
            t.description,
            Money.unparse(t.withdrawn),
            Money.unparse(t.deposited),
            t.category,
            Money.unparse(t.balance),
        ];
    }

    export function eq(a: Transaction, b: Transaction): boolean {
        return Date.eq(a.date, b.date)
            && a.description == b.description
            && Money.eq(a.withdrawn, b.withdrawn)
            && Money.eq(a.deposited, b.deposited)
            && Money.eq(a.balance, b.balance);
    }
}

export function import_data() {
    let spread = SpreadsheetApp.getActiveSpreadsheet();

    let import_sheet = spread.getSheetByName('import');
    assert(import_sheet != null, 'no sheet import');

    let raw_sheet = spread.getSheetByName('raw');
    assert(raw_sheet != null, 'no sheet raw');

    let count = import_sheet.getLastRow() - 1; // 1 row of header, the rest is data
    let data = Transaction.load(import_sheet, 0, count);

    let recent = Transaction.load(raw_sheet, 0);
    let recent_index = data.findIndex(t => Transaction.eq(t, recent));

    console.log(recent);
    for (let x of data)
        console.log(x);
    assert(recent_index != -1, 'no recent :(');

    if (recent_index > 0) {
        data = data.slice(0, recent_index);

        raw_sheet.insertRowsBefore(2, data.length);
        Transaction.save(raw_sheet, 0, data);

        let extra = data.map((v, i) => [
            `=D${2 + i} - C${2 + i}`,
            ``,
            `=TRANSACTION_DESCRIPTION(B${2 + i})`,
        ]);
        raw_sheet.getRange(2, 7, data.length, 3).setValues(extra);

        Logger.log(`import successful: ${data.length}`);

        import_sheet.getRange(2, 1, data.length).setBackground('#d9ead3');
    }
}

export function display_data() {
    let spread = SpreadsheetApp.getActiveSpreadsheet();

    let raw_sheet = spread.getSheetByName('raw');
    assert(raw_sheet != null, 'no sheet raw');

    let display_sheet = spread.getSheetByName('display');
    assert(display_sheet != null, 'no sheet display');

    let categories_sheet = spread.getSheetByName('categories');
    assert(categories_sheet != null, 'no sheet categories');

    let total = raw_sheet.getLastRow() - 1;

    let data: Transaction[];
    if (display_sheet.getLastRow() <= 2) {
        data = Transaction.load(raw_sheet, 0, total);
    } else {
        let rows = display_sheet.getLastRow() - 1;
        let refs = display_sheet.getRange(2, 1, rows, 1)
            .getFormulas()
            .map(row => row[0]);

        let previous = refs.find(i => i !== '');
        assert(previous !== undefined, 'invalid state');

        let previous_ref = /^=raw!A(\d+)$/.exec(previous)?.[1];
        assert(previous_ref !== undefined, 'invalid state');

        let previous_index = parseInt(previous_ref) - 2;
        data = Transaction.load(raw_sheet, 0, previous_index + 1);
        display_sheet.deleteRows(2, 2);
    }

    data.reverse();

    while (data.length > 0) {
        let index = data.findIndex(d => d.date.month != data[0].date.month || d.date.year != data[0].date.year);
        if (index == -1)
            index = data.length;

        let group = data.splice(0, index);

        let values = group.map((t, i) => [
            `=raw!A${2 + data.length + i}`,
            `=raw!G${2 + data.length + i}`,
            `=raw!F${2 + data.length + i}`,
            `=raw!H${2 + data.length + i}`,
            `=raw!I${2 + data.length + i}`,
        ]);

        display_sheet.insertRowsBefore(2, values.length);
        let depth = display_sheet.getRowGroupDepth(2);
        let body = display_sheet.getRange(2, 1, values.length, 5);
        body.setFontWeight(null);
        body.setValues(values);
        body.shiftRowGroupDepth(1 - depth);
        body.setHorizontalAlignments(values.map(() => [
            'left',
            'right' as any,
            'right' as any,
            'left',
            'left',
        ]));

        display_sheet.getRowGroup(2, 1)!.collapse();

        let date = { ...group[0].date, day: 1 };
        let net = `=SUM(B3:B${2 + group.length})`;
        let balance = `=C3`;
        display_sheet.insertRowBefore(2);
        let header = display_sheet.getRange(2, 1, 1, 4);
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

export function TRANSACTION_DESCRIPTION(raw: string) {
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

function assert(condition: boolean, message: string): asserts condition {
    if (!condition) {
        Logger.log(message);
        throw new Error(message);
    }
}

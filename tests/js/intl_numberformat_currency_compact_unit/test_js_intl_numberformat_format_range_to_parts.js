// vybe-test: js/intl_numberformat_currency_compact_unit/test_js_intl_numberformat_format_range_to_parts
// origin: languages/js/tests/js/test_js_intl_numberformat_currency_compact_unit.rs

function __line(...args) {
    // console.log joins its arguments with a single space. String() is the
    // coercion Vybe's logging host applies to each one.
    return args.map(String).join(" ");
}

function __check(got, want) {
    if (got !== want) {
        console.log("FAIL: want [" + want + "] got [" + got + "]");
        throw new Error("assertion failed");
    }
}

const formatter = new Intl.NumberFormat("en-US", { style: "currency", currency: "USD" });
const parts = formatter.formatRangeToParts(10, 20);
__check(__line(parts.some(p => p.type === "currency")), "true");

// vybe-test: js/intl_numberformat_currency_compact_unit/test_js_intl_numberformat_missing_currency_throws_typeerror
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

try {
    new Intl.NumberFormat("en-US", { style: "currency" }); // style: "currency" REQUIRES currency option!
} catch (e) {
    __check(__line("Currency Option Missing TypeError"), "Currency Option Missing TypeError");
}

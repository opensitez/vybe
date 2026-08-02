// vybe-test: js/dataview_get_set_endianness_methods/test_js_dataview_get_set_bigint64_biguint64
// origin: languages/js/tests/js/test_js_dataview_get_set_endianness_methods.rs

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

const buffer = new ArrayBuffer(16);
const dv = new DataView(buffer);
dv.setBigInt64(0, -9007199254740991n, true);
dv.setBigUint64(8, 18446744073709551615n, false);

__check(__line(dv.getBigInt64(0, true).toString() + "|" + dv.getBigUint64(8, false).toString()), "-9007199254740991|18446744073709551615");

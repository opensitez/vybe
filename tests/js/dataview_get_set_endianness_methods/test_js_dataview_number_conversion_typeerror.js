// vybe-test: js/dataview_get_set_endianness_methods/test_js_dataview_number_conversion_typeerror
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

const buffer = new ArrayBuffer(8);
const dv = new DataView(buffer);
try {
    dv.setUint32(0, 100n); // Passing BigInt to Number method throws TypeError!
} catch (e) {
    __check(__line("Number DataView Conversion TypeError"), "Number DataView Conversion TypeError");
}

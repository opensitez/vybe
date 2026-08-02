// vybe-test: js/dataview_get_set_endianness_methods/test_js_dataview_get_set_float16_es2024
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

const buffer = new ArrayBuffer(2);
const dv = new DataView(buffer);
dv.setFloat16(0, 1.5, true);
__check(__line(dv.getFloat16(0, true)), "1.5");

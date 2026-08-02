// vybe-test: js/dataview_get_set_endianness_methods/test_js_dataview_get_set_float32_float64
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

const buffer = new ArrayBuffer(12);
const dv = new DataView(buffer);
dv.setFloat32(0, 3.14, true);
dv.setFloat64(4, 2.718281828459045, false);

__check(__line(dv.getFloat32(0, true).toFixed(2) + "|" + dv.getFloat64(4, false)), "3.14|2.718281828459045");

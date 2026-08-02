// vybe-test: js/dataview_get_set_endianness_methods/test_js_dataview_nan_floating_point_write_read
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

const buffer = new ArrayBuffer(4);
const dv = new DataView(buffer);
dv.setFloat32(0, NaN);
__check(__line(Number.isNaN(dv.getFloat32(0))), "true");

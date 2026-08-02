// vybe-test: js/dataview_get_set_endianness_methods/test_js_dataview_out_of_bounds_read_throws_rangeerror
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
try {
    dv.getInt32(2); // Reads 4 bytes starting at offset 2 (exceeds buffer length 4)!
} catch (e) {
    __check(__line("DataView Read OutOfBounds RangeError"), "DataView Read OutOfBounds RangeError");
}

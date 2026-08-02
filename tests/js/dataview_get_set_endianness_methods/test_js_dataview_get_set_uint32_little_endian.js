// vybe-test: js/dataview_get_set_endianness_methods/test_js_dataview_get_set_uint32_little_endian
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
dv.setUint32(0, 0xDEADBEEF, true); // Little endian write

__check(__line(dv.getUint32(0, true).toString(16).toUpperCase()), "DEADBEEF");

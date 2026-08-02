// vybe-test: js/dataview_get_set_endianness_methods/test_js_dataview_detached_buffer_access_throws_typeerror
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
buffer.transfer();
try {
    dv.getUint8(0);
} catch (e) {
    __check(__line("DataView Detached Buffer TypeError"), "DataView Detached Buffer TypeError");
}

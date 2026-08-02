// vybe-test: js/dataview_get_set_endianness_methods/test_js_dataview_resizable_buffer_auto_byte_length
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

const buffer = new ArrayBuffer(8, { maxByteLength: 32 });
const dv = new DataView(buffer);
__check(__line(dv.byteLength), "8");
buffer.resize(16);
__check(__line(dv.byteLength), "16");

// vybe-test: js/dataview_get_set_endianness_methods/test_js_dataview_byte_length_and_byte_offset_properties
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
const dv = new DataView(buffer, 4, 8);
__check(__line(`${dv.byteOffset}:${dv.byteLength}:${dv.buffer.byteLength}`), "4:8:16");

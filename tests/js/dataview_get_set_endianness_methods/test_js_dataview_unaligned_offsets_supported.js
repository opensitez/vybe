// vybe-test: js/dataview_get_set_endianness_methods/test_js_dataview_unaligned_offsets_supported
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
dv.setUint32(1, 0x12345678, true); // DataView supports unaligned byte offsets!
__check(__line(dv.getUint32(1, true).toString(16)), "12345678");

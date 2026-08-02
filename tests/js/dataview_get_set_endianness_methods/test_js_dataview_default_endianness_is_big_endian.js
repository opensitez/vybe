// vybe-test: js/dataview_get_set_endianness_methods/test_js_dataview_default_endianness_is_big_endian
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
dv.setUint16(0, 0x0102); // Omitted littleEndian flag defaults to FALSE (Big-Endian)
const u8 = new Uint8Array(buffer);
__check(__line(`${u8[0]}:${u8[1]}`), "1:2");

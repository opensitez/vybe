// vybe-test: js/dataview_get_set_endianness_methods/test_js_dataview_get_set_int8_uint8
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
dv.setInt8(0, -128);
dv.setUint8(1, 255);

__check(__line(`${dv.getInt8(0)}:${dv.getUint8(1)}`), "-128:255");

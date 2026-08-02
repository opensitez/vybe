// vybe-test: js/dataview_typed_array_deep/dataview_read_write_uint16_endianness
// origin: languages/js/tests/js/test_dataview_typed_array_deep.rs

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

const buf = new ArrayBuffer(2);
const view = new DataView(buf);
view.setUint16(0, 0x0102, true); // little-endian
__check(__line(view.getUint8(0)), "2"); // low byte
__check(__line(view.getUint8(1)), "1"); // high byte

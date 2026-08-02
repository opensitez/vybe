// vybe-test: js/typed_array_advanced/dataview_little_endian_vs_big_endian
// origin: languages/js/tests/js/test_typed_array_advanced.rs

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

const buf = new ArrayBuffer(4);
const dv = new DataView(buf);
dv.setUint16(0, 0x0102, true);  // little endian
__check(__line(dv.getUint8(0)), "2");   // low byte first
__check(__line(dv.getUint8(1)), "1");   // high byte

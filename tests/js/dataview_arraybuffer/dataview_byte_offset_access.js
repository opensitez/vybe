// vybe-test: js/dataview_arraybuffer/dataview_byte_offset_access
// origin: languages/js/tests/js/test_dataview_arraybuffer.rs

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

const buf = new ArrayBuffer(8);
const dv = new DataView(buf);
dv.setUint8(0, 10);
dv.setUint8(1, 20);
dv.setUint8(2, 30);
__check(__line(dv.getUint8(1)), "20");

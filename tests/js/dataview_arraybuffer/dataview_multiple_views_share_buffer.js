// vybe-test: js/dataview_arraybuffer/dataview_multiple_views_share_buffer
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

const buf = new ArrayBuffer(4);
const dv = new DataView(buf);
const u8 = new Uint8Array(buf);
dv.setUint8(0, 42);
__check(__line(u8[0]), "42");

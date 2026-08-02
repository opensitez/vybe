// vybe-test: js/dataview_arraybuffer/dataview_set_and_get_int8
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
dv.setInt8(0, -128);
__check(__line(dv.getInt8(0)), "-128");

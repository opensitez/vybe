// vybe-test: js/dataview_arraybuffer/dataview_set_and_get_float64
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

const buf = new ArrayBuffer(16);
const dv = new DataView(buf);
dv.setFloat64(0, 3.141592653589793);
const val = dv.getFloat64(0);
__check(__line(val.toFixed(6)), "3.141593");

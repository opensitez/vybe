// vybe-test: js/typed_arrays_deep/dataview_read_write
// origin: languages/js/tests/js/test_typed_arrays_deep.rs

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
const view = new DataView(buf);
view.setFloat64(0, Math.PI, false);  // big endian
const pi = view.getFloat64(0, false);
__check(__line(Math.abs(pi - Math.PI) < 1e-10), "true");
view.setInt16(0, -1000, true);  // little endian
__check(__line(view.getInt16(0, true)), "-1000");

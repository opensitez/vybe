// vybe-test: js/dataview_typed_array_deep/dataview_read_write_int8
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

const buf = new ArrayBuffer(4);
const view = new DataView(buf);
view.setInt8(0, -1);
view.setInt8(1, 127);
__check(__line(view.getInt8(0)), "-1");
__check(__line(view.getInt8(1)), "127");

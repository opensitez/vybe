// vybe-test: js/shared_array_buffer_view_sharing/test_js_shared_array_buffer_dataview_access
// origin: languages/js/tests/js/test_js_shared_array_buffer_view_sharing.rs

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

const sab = new SharedArrayBuffer(8);
const dv = new DataView(sab);
dv.setInt32(0, 12345678, true);
__check(__line(dv.getInt32(0, true)), "12345678");

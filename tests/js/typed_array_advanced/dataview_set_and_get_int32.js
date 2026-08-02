// vybe-test: js/typed_array_advanced/dataview_set_and_get_int32
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
dv.setUint32(0, 0xDEADBEEF);
__check(__line(dv.getUint32(0).toString(16)), "deadbeef");

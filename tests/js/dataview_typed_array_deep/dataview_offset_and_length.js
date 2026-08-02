// vybe-test: js/dataview_typed_array_deep/dataview_offset_and_length
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

const buf = new ArrayBuffer(10);
const view = new DataView(buf, 2, 4); // offset 2, length 4
__check(__line(view.byteOffset), "2");
__check(__line(view.byteLength), "4");
view.setInt8(0, 42); // relative to offset
__check(__line(view.getInt8(0)), "42");

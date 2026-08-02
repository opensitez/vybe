// vybe-test: js/structured_clone_typed_arrays_array_buffers/test_js_structured_clone_dataview_byte_offset_preserved
// origin: languages/js/tests/js/test_js_structured_clone_typed_arrays_array_buffers.rs

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
const dv = new DataView(buf, 4, 8);
const clone = structuredClone(dv);
__check(__line(clone.byteOffset + "|" + clone.byteLength), "0|8");

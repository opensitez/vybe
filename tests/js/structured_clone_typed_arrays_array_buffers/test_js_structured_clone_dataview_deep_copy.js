// vybe-test: js/structured_clone_typed_arrays_array_buffers/test_js_structured_clone_dataview_deep_copy
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

const buf = new ArrayBuffer(8);
const dv = new DataView(buf, 2, 4);
dv.setInt16(0, 1234, true);

const cloneDV = structuredClone(dv);
__check(__line((cloneDV instanceof DataView) + "|" + (cloneDV.buffer !== buf) + "|" + cloneDV.getInt16(0, true)), "true|true|1234");

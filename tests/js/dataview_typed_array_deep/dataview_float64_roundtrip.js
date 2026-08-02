// vybe-test: js/dataview_typed_array_deep/dataview_float64_roundtrip
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

const buf = new ArrayBuffer(8);
const view = new DataView(buf);
const pi = Math.PI;
view.setFloat64(0, pi);
__check(__line(view.getFloat64(0) === pi), "true");

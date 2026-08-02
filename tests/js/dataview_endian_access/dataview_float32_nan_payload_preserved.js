// vybe-test: js/dataview_endian_access/dataview_float32_nan_payload_preserved
// origin: languages/js/tests/js/test_dataview_endian_access.rs

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

const v=new DataView(new ArrayBuffer(4)); v.setFloat32(0, NaN, true); __check(__line(Number.isNaN(v.getFloat32(0,true))), "true");

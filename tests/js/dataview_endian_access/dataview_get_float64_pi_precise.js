// vybe-test: js/dataview_endian_access/dataview_get_float64_pi_precise
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

const b=new ArrayBuffer(8); const v=new DataView(b); v.setFloat64(0, Math.PI, true); __check(__line(v.getFloat64(0, true)>3.14), "true");

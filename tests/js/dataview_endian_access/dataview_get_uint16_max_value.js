// vybe-test: js/dataview_endian_access/dataview_get_uint16_max_value
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

const b=new ArrayBuffer(2); const v=new DataView(b); v.setUint16(0, 0xffff, true); __check(__line(v.getUint16(0, true)), "65535");

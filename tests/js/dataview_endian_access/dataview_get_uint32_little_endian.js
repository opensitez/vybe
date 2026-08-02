// vybe-test: js/dataview_endian_access/dataview_get_uint32_little_endian
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

const b=new ArrayBuffer(4); const v=new DataView(b); v.setUint32(0, 0x11223344, true); __check(__line(v.getUint32(0, true).toString(16)), "11223344");

// vybe-test: js/dataview_endian_access/dataview_get_uint8_reads_unsigned_byte
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

const b=new ArrayBuffer(1); const v=new DataView(b); v.setUint8(0,200); __check(__line(v.getUint8(0)), "200");

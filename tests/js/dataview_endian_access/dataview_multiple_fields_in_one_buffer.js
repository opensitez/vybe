// vybe-test: js/dataview_endian_access/dataview_multiple_fields_in_one_buffer
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

const v=new DataView(new ArrayBuffer(6)); v.setUint16(0,1,true); v.setUint32(2,2,true); __check(__line(v.getUint16(0,true)), "1");__check(__line(v.getUint32(2,true)), "2");

// vybe-test: js/dataview_endian_access/dataview_get_int32_on_shared_buffer
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

const sab=new SharedArrayBuffer(4); const v=new DataView(sab); v.setInt32(0,99,true); __check(__line(v.getInt32(0,true)), "99");

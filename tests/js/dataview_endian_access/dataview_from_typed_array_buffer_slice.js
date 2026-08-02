// vybe-test: js/dataview_endian_access/dataview_from_typed_array_buffer_slice
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

const arr=new Uint8Array([1,2,3,4]); const v=new DataView(arr.buffer,1,2); __check(__line(v.getUint8(0)), "2");__check(__line(v.getUint8(1)), "3");

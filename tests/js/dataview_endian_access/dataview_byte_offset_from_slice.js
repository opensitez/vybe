// vybe-test: js/dataview_endian_access/dataview_byte_offset_from_slice
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

const v=new DataView(new ArrayBuffer(8), 2, 4); __check(__line(v.byteOffset), "2");__check(__line(v.byteLength), "4");

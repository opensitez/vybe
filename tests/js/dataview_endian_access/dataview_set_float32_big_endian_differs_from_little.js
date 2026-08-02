// vybe-test: js/dataview_endian_access/dataview_set_float32_big_endian_differs_from_little
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

const v=new DataView(new ArrayBuffer(4)); v.setFloat32(0,1,true); const le=v.getFloat32(0,true); const be=v.getFloat32(0,false); __check(__line(le!==be), "true");

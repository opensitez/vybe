// vybe-test: js/dataview_endian_access/dataview_set_out_of_range_throws_range_error
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

const v=new DataView(new ArrayBuffer(1)); try{v.setInt32(0,1);}catch(e){__check(__line(e instanceof RangeError), "true");}

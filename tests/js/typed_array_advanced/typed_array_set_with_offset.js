// vybe-test: js/typed_array_advanced/typed_array_set_with_offset
// origin: languages/js/tests/js/test_typed_array_advanced.rs

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

const ta = new Uint8Array(5);
ta.set([1, 2], 2); // start at index 2
__check(__line(ta.join(",")), "0,0,1,2,0");

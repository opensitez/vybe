// vybe-test: js/typed_array_from_of_constructors/test_js_typedarray_of_element_coercion
// origin: languages/js/tests/js/test_js_typed_array_from_of_constructors.rs

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

const u8 = Uint8Array.of("50", 256, true); // "50"->50, 256->0, true->1
__check(__line(u8.join(",")), "50,0,1");

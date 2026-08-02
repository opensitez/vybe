// vybe-test: js/object_has_own_vs_has_own_property/test_js_object_has_own_typed_array_indices
// origin: languages/js/tests/js/test_js_object_has_own_vs_has_own_property.rs

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

const u8 = new Uint8Array([5, 10]);
__check(__line(`${Object.hasOwn(u8, 0)}:${Object.hasOwn(u8, 1)}:${Object.hasOwn(u8, 2)}`), "true:true:false");

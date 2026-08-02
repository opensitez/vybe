// vybe-test: js/object_has_own_vs_has_own_property/test_js_object_has_own_array_indices_and_length
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

const arr = [10, 20];
__check(__line(`${Object.hasOwn(arr, 0)}:${Object.hasOwn(arr, 1)}:${Object.hasOwn(arr, 2)}:${Object.hasOwn(arr, "length")}`), "true:true:false:true");

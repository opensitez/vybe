// vybe-test: js/typed_array_from_of_constructors/test_js_typedarray_from_array_factory
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

const i32 = Int32Array.from([100, 200, 300]);
__check(__line(i32.length + "|" + i32.join(",")), "3|100,200,300");

// vybe-test: js/typed_array_from_of_constructors/test_js_typedarray_from_non_callable_mapfn_throws_typeerror
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

try {
    Uint8Array.from([1, 2], "not_a_function");
} catch (e) {
    __check(__line("Uint8Array.from Non-Callable MapFn TypeError"), "Uint8Array.from Non-Callable MapFn TypeError");
}

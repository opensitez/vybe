// vybe-test: js/nested_destructuring_array_object_mix/test_js_nested_destructuring_undefined_nested_property_throws
// origin: languages/js/tests/js/test_js_nested_destructuring_array_object_mix.rs

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

const obj = {};
try {
    const { nested: { val } } = obj;
} catch (e) {
    __check(__line("Nested Undefined TypeError"), "Nested Undefined TypeError");
}

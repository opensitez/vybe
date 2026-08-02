// vybe-test: js/nullish_coalescing_and_optional_chaining_combinations/test_js_optional_chaining_call_on_non_function_throws_typeerror
// origin: languages/js/tests/js/test_js_nullish_coalescing_and_optional_chaining_combinations.rs

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

const obj = { fn: 123 };
try {
    obj.fn?.();
} catch (e) {
    __check(__line(e.name), "TypeError");
}

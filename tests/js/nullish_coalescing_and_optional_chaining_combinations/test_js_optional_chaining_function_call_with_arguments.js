// vybe-test: js/nullish_coalescing_and_optional_chaining_combinations/test_js_optional_chaining_function_call_with_arguments
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

const adder = (a, b) => a + b;
const missing = null;
__check(__line(adder?.(10, 20) + "|" + (missing?.(10, 20) === undefined)), "30|true");

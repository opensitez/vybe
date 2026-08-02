// vybe-test: js/function_call_apply_arguments_array/test_js_function_apply_array_like_arguments
// origin: languages/js/tests/js/test_js_function_call_apply_arguments_array.rs

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

function sum(a, b, c) {
    return a + b + c;
}
const arrayLike = { 0: 10, 1: 20, 2: 30, length: 3 };
__check(__line(sum.apply(null, arrayLike)), "60");

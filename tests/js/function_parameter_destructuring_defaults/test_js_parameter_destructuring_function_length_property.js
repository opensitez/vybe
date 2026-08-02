// vybe-test: js/function_parameter_destructuring_defaults/test_js_parameter_destructuring_function_length_property
// origin: languages/js/tests/js/test_js_function_parameter_destructuring_defaults.rs

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

function f1(a, { b }, c = 1) {}
function f2({ a }, b = 2, c) {}
__check(__line(f1.length + "|" + f2.length), "1|1"); // Parameters up to first default/destructured without outer fallback

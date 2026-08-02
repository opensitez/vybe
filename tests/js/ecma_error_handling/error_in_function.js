// vybe-test: js/ecma_error_handling/error_in_function
// origin: languages/js/tests/js/test_ecma_error_handling.rs

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

function divide(a, b) {
    if (b === 0) throw new Error("Division by zero");
    return a / b;
}
try {
    console.log(divide(10, 2));
    console.log(divide(10, 0));
} catch (e) {
    console.log(e.message);
}

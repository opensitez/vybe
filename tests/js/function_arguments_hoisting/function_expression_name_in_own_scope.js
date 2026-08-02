// vybe-test: js/function_arguments_hoisting/function_expression_name_in_own_scope
// origin: languages/js/tests/js/test_function_arguments_hoisting.rs

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

const factorial = function fact(n) {
    return n <= 1 ? 1 : n * fact(n - 1);
};
__check(__line(factorial(5)), "120");
__check(__line(typeof fact), "undefined"); // not accessible outside

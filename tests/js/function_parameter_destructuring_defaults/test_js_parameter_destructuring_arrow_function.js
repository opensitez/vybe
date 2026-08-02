// vybe-test: js/function_parameter_destructuring_defaults/test_js_parameter_destructuring_arrow_function
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

const getFull = ({ first, last }) => `${first} ${last}`;
__check(__line(getFull({ first: "John", last: "Doe" })), "John Doe");

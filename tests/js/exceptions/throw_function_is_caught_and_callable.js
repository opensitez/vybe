// vybe-test: js/exceptions/throw_function_is_caught_and_callable
// origin: languages/js/tests/js/test_exceptions.rs

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

let res = "";
try {
    throw function() { return "called"; };
} catch (e) {
    res = e();
}
__check(__line(res), "called");

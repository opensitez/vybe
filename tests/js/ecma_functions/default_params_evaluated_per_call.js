// vybe-test: js/ecma_functions/default_params_evaluated_per_call
// origin: languages/js/tests/js/test_ecma_functions.rs

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

let n = 0;
function next(value = ++n) {
    console.log(value);
}
next();
next();
next(10);
next();

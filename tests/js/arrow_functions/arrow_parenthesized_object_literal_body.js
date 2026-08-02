// vybe-test: js/arrow_functions/arrow_parenthesized_object_literal_body
// origin: languages/js/tests/js/test_arrow_functions.rs

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

const mk=(x)=>({k:x}); __check(__line(mk(3).k), "3");

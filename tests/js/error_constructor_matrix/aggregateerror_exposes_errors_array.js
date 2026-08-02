// vybe-test: js/error_constructor_matrix/aggregateerror_exposes_errors_array
// origin: languages/js/tests/js/test_error_constructor_matrix.rs

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

const e = new AggregateError([1, 2, 3], "many");
__check(__line(e.errors.length), "3");
__check(__line(e.message), "many");

// vybe-test: js/error_cause_chain_unwinding/test_js_error_cause_heterogeneous_types_in_chain
// origin: languages/js/tests/js/test_js_error_cause_chain_unwinding.rs

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

const numCause = 404;
const strCause = "NOT_FOUND";
const err1 = new Error("HttpError", { cause: numCause });
const err2 = new Error("ApiError", { cause: strCause });

__check(__line(`${typeof err1.cause}:${err1.cause} | ${typeof err2.cause}:${err2.cause}`), "number:404 | string:NOT_FOUND");

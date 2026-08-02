// vybe-test: js/error_cause_aggregate/error_cause_works_on_all_error_subtypes
// origin: languages/js/tests/js/test_error_cause_aggregate.rs

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

const cause = new Error("root cause");
const types = [TypeError, RangeError, ReferenceError, SyntaxError, URIError, EvalError];
for (const T of types) {
    const e = new T("wrapper", { cause });
    console.log(e.cause === cause);
}

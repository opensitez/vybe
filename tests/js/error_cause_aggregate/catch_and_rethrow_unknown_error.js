// vybe-test: js/error_cause_aggregate/catch_and_rethrow_unknown_error
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

class DatabaseError extends Error {}
let result = "";
try {
    try { throw new TypeError("unexpected type"); }
    catch (e) {
        if (e instanceof DatabaseError) result = "db";
        else throw e;
    }
} catch (e) {
    result = "rethrown:" + e.constructor.name;
}
__check(__line(result), "rethrown:TypeError");

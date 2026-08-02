// vybe-test: js/error_types_deep/catch_by_type_in_hierarchy
// origin: languages/js/tests/js/test_error_types_deep.rs

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

function risky(x) {
    if (x < 0) throw new RangeError("negative");
    if (typeof x !== "number") throw new TypeError("not a number");
}
function handle(x) {
    try {
        risky(x);
    } catch (e) {
        if (e instanceof RangeError) return "range";
        if (e instanceof TypeError) return "type";
        throw e;
    }
}
__check(__line(handle(-1)), "range");
__check(__line(handle("foo")), "type");

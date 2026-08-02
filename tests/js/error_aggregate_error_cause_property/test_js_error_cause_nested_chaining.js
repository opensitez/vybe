// vybe-test: js/error_aggregate_error_cause_property/test_js_error_cause_nested_chaining
// origin: languages/js/tests/js/test_js_error_aggregate_error_cause_property.rs

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

const e1 = new Error("Level 1");
const e2 = new Error("Level 2", { cause: e1 });
const e3 = new Error("Level 3", { cause: e2 });
__check(__line(`${e3.message} -> ${e3.cause.message} -> ${e3.cause.cause.message}`), "Level 3 -> Level 2 -> Level 1");

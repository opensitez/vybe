// vybe-test: js/structured_clone_error_types_serialization/test_js_structured_clone_error_with_cyclic_cause
// origin: languages/js/tests/js/test_js_structured_clone_error_types_serialization.rs

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

const e1 = new Error("E1");
const e2 = new Error("E2", { cause: e1 });
e1.cause = e2;

const clone = structuredClone(e2);
__check(__line((clone.cause.cause === clone)), "true");

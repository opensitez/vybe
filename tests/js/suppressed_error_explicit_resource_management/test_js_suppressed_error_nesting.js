// vybe-test: js/suppressed_error_explicit_resource_management/test_js_suppressed_error_nesting
// origin: languages/js/tests/js/test_js_suppressed_error_explicit_resource_management.rs

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
const e2 = new Error("E2");
const e3 = new Error("E3");
const inner = new SuppressedError(e1, e2);
const outer = new SuppressedError(inner, e3);

__check(__line(outer.error.error.message + " -> " + outer.error.suppressed.message), "E1 -> E2");

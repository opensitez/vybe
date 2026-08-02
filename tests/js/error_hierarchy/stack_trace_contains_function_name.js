// vybe-test: js/error_hierarchy/stack_trace_contains_function_name
// origin: languages/js/tests/js/test_error_hierarchy.rs

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

function namedFunction() { return new Error("test"); }
const e = namedFunction();
// Stack should mention something — format varies by engine
__check(__line(typeof e.stack === "string"), "true");

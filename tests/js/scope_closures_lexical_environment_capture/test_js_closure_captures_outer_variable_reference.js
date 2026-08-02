// vybe-test: js/scope_closures_lexical_environment_capture/test_js_closure_captures_outer_variable_reference
// origin: languages/js/tests/js/test_js_scope_closures_lexical_environment_capture.rs

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

function createCounter() {
    let count = 0;
    return () => ++count;
}
const counter = createCounter();
__check(__line(`${counter()}:${counter()}:${counter()}`), "1:2:3");

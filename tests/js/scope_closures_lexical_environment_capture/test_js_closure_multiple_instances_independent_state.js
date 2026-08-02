// vybe-test: js/scope_closures_lexical_environment_capture/test_js_closure_multiple_instances_independent_state
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

function makeAdder(x) {
    return (y) => x + y;
}
const add5 = makeAdder(5);
const add10 = makeAdder(10);
__check(__line(add5(3) + "|" + add10(3)), "8|13");

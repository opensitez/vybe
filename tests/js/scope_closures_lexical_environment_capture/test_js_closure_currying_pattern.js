// vybe-test: js/scope_closures_lexical_environment_capture/test_js_closure_currying_pattern
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

const curry3 = (fn) => (a) => (b) => (c) => fn(a, b, c);
const sum = (a, b, c) => a + b + c;
__check(__line(curry3(sum)(1)(2)(3)), "6");

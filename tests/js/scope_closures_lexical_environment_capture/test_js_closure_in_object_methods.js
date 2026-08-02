// vybe-test: js/scope_closures_lexical_environment_capture/test_js_closure_in_object_methods
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

const module = (() => {
    let internalState = 0;
    return {
        increment() { internalState++; },
        read() { return internalState; }
    };
})();
module.increment();
module.increment();
__check(__line(module.read()), "2");

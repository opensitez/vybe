// vybe-test: js/scope_closures_lexical_environment_capture/test_js_closure_loop_var_shared_binding_issue
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

const funcs = [];
for (var i = 0; i < 3; i++) {
    funcs.push(() => i);
}
console.log(funcs.map(f => f()).join(",")); // var shares single binding -> returns 3,3,3

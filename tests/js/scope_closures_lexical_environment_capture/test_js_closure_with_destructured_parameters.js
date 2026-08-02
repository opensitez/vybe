// vybe-test: js/scope_closures_lexical_environment_capture/test_js_closure_with_destructured_parameters
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

function process({ name, age }) {
    return () => `${name} is ${age}`;
}
const desc = process({ name: "Bob", age: 25 });
__check(__line(desc()), "Bob is 25");

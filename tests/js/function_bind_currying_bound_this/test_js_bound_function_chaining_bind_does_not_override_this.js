// vybe-test: js/function_bind_currying_bound_this/test_js_bound_function_chaining_bind_does_not_override_this
// origin: languages/js/tests/js/test_js_function_bind_currying_bound_this.rs

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

function getThisName() {
    return this.name;
}
const obj1 = { name: "First" };
const obj2 = { name: "Second" };

const bound1 = getThisName.bind(obj1);
const bound2 = bound1.bind(obj2); // Re-binding does NOT change the initial 'this' binding!
__check(__line(bound2()), "First");

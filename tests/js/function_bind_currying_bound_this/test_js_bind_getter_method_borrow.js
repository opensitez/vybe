// vybe-test: js/function_bind_currying_bound_this/test_js_bind_getter_method_borrow
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

const obj = { get val() { return this._v; } };
const getter = Object.getOwnPropertyDescriptor(obj, "val").get.bind({ _v: 42 });
__check(__line(getter()), "42");

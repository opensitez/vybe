// vybe-test: js/function_bind_currying_bound_this/test_js_bound_arrow_function_ignores_bound_this
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

const arrow = () => this;
const bound = arrow.bind({ a: 1 });
__check(__line(bound() === this), "true"); // Arrow function 'this' is lexically static, bind thisArg is ignored!

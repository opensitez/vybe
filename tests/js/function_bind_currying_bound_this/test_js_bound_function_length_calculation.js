// vybe-test: js/function_bind_currying_bound_this/test_js_bound_function_length_calculation
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

function sum(a, b, c, d) {}
const bound1 = sum.bind(null, 1);
const bound2 = sum.bind(null, 1, 2);
__check(__line(`${sum.length}:${bound1.length}:${bound2.length}`), "4:3:2");

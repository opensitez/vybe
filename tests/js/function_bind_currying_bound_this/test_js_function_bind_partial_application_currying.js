// vybe-test: js/function_bind_currying_bound_this/test_js_function_bind_partial_application_currying
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

function add(a, b, c) {
    return a + b + c;
}
const add5 = add.bind(null, 5);
const add5And10 = add5.bind(null, 10);
__check(__line(add5(2, 3) + "|" + add5And10(4)), "10|19");

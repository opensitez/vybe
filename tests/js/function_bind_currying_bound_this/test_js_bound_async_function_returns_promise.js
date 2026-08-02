// vybe-test: js/function_bind_currying_bound_this/test_js_bound_async_function_returns_promise
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

async function asyncFn(a) {
    return a + this.val;
}
const boundAsync = asyncFn.bind({ val: 5 }, 10);
(async () => {
    console.log(await boundAsync());
})();

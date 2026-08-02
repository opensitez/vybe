// vybe-test: js/proxy_apply_construct_invocations/test_js_proxy_apply_trap_function_call
// origin: languages/js/tests/js/test_js_proxy_apply_construct_invocations.rs

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

function sum(a, b) { return a + b; }
const proxy = new Proxy(sum, {
    apply(target, thisArg, args) {
        return target(...args) * 10;
    }
});
__check(__line(proxy(2, 3)), "50");

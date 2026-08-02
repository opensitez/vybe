// vybe-test: js/proxy_apply_construct_invocations/test_js_proxy_apply_trap_with_reflect_apply
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

function multiply(x, y) { return x * y; }
const proxy = new Proxy(multiply, {
    apply(target, thisArg, args) {
        return Reflect.apply(target, thisArg, [args[0] + 1, args[1] + 1]);
    }
});
__check(__line(proxy(2, 3)), "12"); // (2+1) * (3+1) = 12

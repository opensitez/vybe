// vybe-test: js/proxy_apply_construct_invocations/test_js_proxy_apply_trap_memoization
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

const cache = new Map();
function fib(n) {
    if (n <= 1) return n;
    return fib(n - 1) + fib(n - 2);
}
const memoFib = new Proxy(fib, {
    apply(target, thisArg, args) {
        const key = args[0];
        if (cache.has(key)) return cache.get(key);
        const res = target(...args);
        cache.set(key, res);
        return res;
    }
});
__check(__line(memoFib(10)), "55");
__check(__line(cache.has(10)), "true");

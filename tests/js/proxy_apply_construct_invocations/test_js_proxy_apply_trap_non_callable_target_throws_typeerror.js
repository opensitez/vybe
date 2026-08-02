// vybe-test: js/proxy_apply_construct_invocations/test_js_proxy_apply_trap_non_callable_target_throws_typeerror
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

const nonFn = { a: 1 };
try {
    const proxy = new Proxy(nonFn, {
        apply(target, thisArg, args) { return 0; }
    });
    proxy();
} catch (e) {
    __check(__line("Non-Callable Apply Error"), "Non-Callable Apply Error");
}

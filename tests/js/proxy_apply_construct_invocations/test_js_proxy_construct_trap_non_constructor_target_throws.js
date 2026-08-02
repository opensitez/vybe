// vybe-test: js/proxy_apply_construct_invocations/test_js_proxy_construct_trap_non_constructor_target_throws
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

const arrowFn = () => {};
try {
    const proxy = new Proxy(arrowFn, {
        construct(target, args) { return {}; }
    });
    new proxy();
} catch (e) {
    __check(__line("Non-Constructor Proxy Error"), "Non-Constructor Proxy Error");
}

// vybe-test: js/proxy_apply_construct_invocations/test_js_proxy_apply_trap_variadic_arguments_forwarding
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

function concatAll(...strings) { return strings.join("-"); }
const proxy = new Proxy(concatAll, {
    apply(target, thisArg, args) {
        return target(...args).toUpperCase();
    }
});
__check(__line(proxy("a", "b", "c")), "A-B-C");

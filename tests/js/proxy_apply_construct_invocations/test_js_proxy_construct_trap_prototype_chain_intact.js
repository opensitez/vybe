// vybe-test: js/proxy_apply_construct_invocations/test_js_proxy_construct_trap_prototype_chain_intact
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

class Widget {}
const ProxyWidget = new Proxy(Widget, {
    construct(target, args, newTarget) {
        return Reflect.construct(target, args, newTarget);
    }
});
const w = new ProxyWidget();
__check(__line(w instanceof Widget), "true");

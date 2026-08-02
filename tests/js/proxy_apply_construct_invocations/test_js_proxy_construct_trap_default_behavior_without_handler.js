// vybe-test: js/proxy_apply_construct_invocations/test_js_proxy_construct_trap_default_behavior_without_handler
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

class Config {
    constructor(v) { this.v = v; }
}
const ProxyConfig = new Proxy(Config, {});
const c = new ProxyConfig(42);
__check(__line(c.v), "42");

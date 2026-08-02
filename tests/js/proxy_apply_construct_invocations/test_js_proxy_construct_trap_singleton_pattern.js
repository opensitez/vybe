// vybe-test: js/proxy_apply_construct_invocations/test_js_proxy_construct_trap_singleton_pattern
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

let instance = null;
class Service {}
const SingletonService = new Proxy(Service, {
    construct(target, args) {
        if (!instance) instance = new target(...args);
        return instance;
    }
});
const s1 = new SingletonService();
const s2 = new SingletonService();
__check(__line(s1 === s2), "true");

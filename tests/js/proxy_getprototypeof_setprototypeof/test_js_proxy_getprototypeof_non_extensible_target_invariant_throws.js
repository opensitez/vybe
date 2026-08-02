// vybe-test: js/proxy_getprototypeof_setprototypeof/test_js_proxy_getprototypeof_non_extensible_target_invariant_throws
// origin: languages/js/tests/js/test_js_proxy_getprototypeof_setprototypeof.rs

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

const target = Object.preventExtensions({ a: 1 });
const actualProto = Object.getPrototypeOf(target);
const wrongProto = {};

const proxy = new Proxy(target, {
    getPrototypeOf() {
        return wrongProto; // Invariant: If target is non-extensible, getPrototypeOf must return target's actual prototype!
    }
});
try {
    Object.getPrototypeOf(proxy);
} catch (e) {
    console.log("getPrototypeOf Non-Extensible Invariant TypeError");
}

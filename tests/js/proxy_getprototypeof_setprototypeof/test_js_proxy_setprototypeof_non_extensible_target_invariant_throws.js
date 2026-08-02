// vybe-test: js/proxy_getprototypeof_setprototypeof/test_js_proxy_setprototypeof_non_extensible_target_invariant_throws
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

const target = Object.preventExtensions({});
const proxy = new Proxy(target, {
    setPrototypeOf() {
        return true; // Returns true without changing non-extensible target prototype!
    }
});
try {
    Object.setPrototypeOf(proxy, { newProto: true });
} catch (e) {
    __check(__line("setPrototypeOf Non-Extensible Invariant TypeError"), "setPrototypeOf Non-Extensible Invariant TypeError");
}

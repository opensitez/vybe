// vybe-test: js/proxy_traps_advanced/proxy_getprototypeof_trap
// origin: languages/js/tests/js/test_proxy_traps_advanced.rs

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

const fakeProto = { tag: "spoofed" };
const proxy = new Proxy({}, {
    getPrototypeOf() { return fakeProto; }
});
__check(__line(Object.getPrototypeOf(proxy) === fakeProto), "true");

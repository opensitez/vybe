// vybe-test: js/proxy_getprototypeof_setprototypeof/test_js_proxy_setprototypeof_cycle_detection_throws_typeerror
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

const target = {};
const proxy = new Proxy(target, {
    setPrototypeOf(t, proto) {
        return Reflect.setPrototypeOf(t, proto);
    }
});
try {
    Object.setPrototypeOf(target, proxy); // Cyclic prototype chain assignment throws TypeError!
} catch (e) {
    __check(__line("Cyclic Prototype Assignment TypeError"), "Cyclic Prototype Assignment TypeError");
}

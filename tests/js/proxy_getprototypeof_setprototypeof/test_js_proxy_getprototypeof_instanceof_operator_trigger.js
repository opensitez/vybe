// vybe-test: js/proxy_getprototypeof_setprototypeof/test_js_proxy_getprototypeof_instanceof_operator_trigger
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

class CustomType {}
const proxy = new Proxy({}, {
    getPrototypeOf() { return CustomType.prototype; }
});
__check(__line(proxy instanceof CustomType), "true");

// vybe-test: js/proxy_getprototypeof_setprototypeof/test_js_proxy_setprototypeof_trap_this_binding
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

let trapThis;
const handler = {
    setPrototypeOf(t, proto) {
        trapThis = this;
        return Reflect.setPrototypeOf(t, proto);
    }
};
const proxy = new Proxy({}, handler);
Object.setPrototypeOf(proxy, {});
__check(__line(trapThis === handler), "true");

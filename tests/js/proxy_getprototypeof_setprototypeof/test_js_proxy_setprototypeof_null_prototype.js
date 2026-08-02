// vybe-test: js/proxy_getprototypeof_setprototypeof/test_js_proxy_setprototypeof_null_prototype
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

const target = { a: 1 };
const proxy = new Proxy(target, {});
Object.setPrototypeOf(proxy, null);
__check(__line(Object.getPrototypeOf(proxy) === null), "true");

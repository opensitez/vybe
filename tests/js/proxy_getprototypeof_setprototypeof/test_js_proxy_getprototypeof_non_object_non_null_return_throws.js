// vybe-test: js/proxy_getprototypeof_setprototypeof/test_js_proxy_getprototypeof_non_object_non_null_return_throws
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

const proxy = new Proxy({}, {
    getPrototypeOf() { return "not_an_object"; }
});
try {
    Object.getPrototypeOf(proxy);
} catch (e) {
    __check(__line("getPrototypeOf Non-Object Return TypeError"), "getPrototypeOf Non-Object Return TypeError");
}

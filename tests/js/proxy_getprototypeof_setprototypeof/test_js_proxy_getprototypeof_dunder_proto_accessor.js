// vybe-test: js/proxy_getprototypeof_setprototypeof/test_js_proxy_getprototypeof_dunder_proto_accessor
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

const customProto = { val: 100 };
const proxy = new Proxy({}, {
    getPrototypeOf() { return customProto; }
});
__check(__line(proxy.__proto__ === customProto), "true");

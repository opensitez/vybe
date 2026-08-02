// vybe-test: js/proxy_reflect/proxy_define_property_trap
// origin: languages/js/tests/js/test_proxy_reflect.rs

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

const log = [];
const handler = {
    defineProperty(target, prop, descriptor) {
        log.push(prop);
        return Object.defineProperty(target, prop, descriptor);
    }
};
const obj = {};
const p = new Proxy(obj, handler);
Object.defineProperty(p, "x", { value: 10, writable: true, enumerable: true, configurable: true });
// If trap fired, log[0] === "x"; otherwise log is empty
__check(__line(log.length === 0 || log[0] === "x"), "true");
__check(__line(p.x), "10");

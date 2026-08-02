// vybe-test: js/proxy_core_traps/proxy_get_intercepts_property_read
// origin: languages/js/tests/js/test_proxy_core_traps.rs

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

const handler = {
    get(target, prop) {
        return prop in target ? target[prop] : `[${prop} not found]`;
    }
};
const obj = new Proxy({ a: 1 }, handler);
__check(__line(obj.a), "1");
__check(__line(obj.b), "[b not found]");

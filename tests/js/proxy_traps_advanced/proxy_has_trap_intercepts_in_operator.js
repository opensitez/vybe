// vybe-test: js/proxy_traps_advanced/proxy_has_trap_intercepts_in_operator
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

const handler = {
    has(target, prop) {
        console.log("has:" + prop);
        return prop in target;
    }
};
const proxy = new Proxy({ a: 1 }, handler);
console.log("a" in proxy);
console.log("b" in proxy);

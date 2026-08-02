// vybe-test: js/proxy_traps_advanced/proxy_delete_property_trap_intercepts
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

const log = [];
const proxy = new Proxy({ a: 1, b: 2 }, {
    deleteProperty(target, prop) {
        log.push("delete:" + prop);
        return delete target[prop];
    }
});
delete proxy.a;
__check(__line(log.join(",")), "delete:a");
__check(__line("a" in proxy), "false");

// vybe-test: js/proxy_core_traps/proxy_delete_property_trap
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

const deleted = [];
const handler = {
    deleteProperty(target, prop) {
        deleted.push(prop);
        return delete target[prop];
    }
};
const obj = new Proxy({ a: 1, b: 2 }, handler);
delete obj.a;
__check(__line(deleted.join(",")), "a");
__check(__line("a" in obj), "false");

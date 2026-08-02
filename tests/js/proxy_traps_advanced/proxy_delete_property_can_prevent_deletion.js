// vybe-test: js/proxy_traps_advanced/proxy_delete_property_can_prevent_deletion
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

const protected_props = new Set(["core"]);
const proxy = new Proxy({ core: 1, temp: 2 }, {
    deleteProperty(target, prop) {
        if (protected_props.has(prop)) return false;
        return delete target[prop];
    }
});
delete proxy.temp;
delete proxy.core;
__check(__line("temp" in proxy), "false");
__check(__line("core" in proxy), "true");

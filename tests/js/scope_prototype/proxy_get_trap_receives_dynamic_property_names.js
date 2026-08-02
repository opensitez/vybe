// vybe-test: js/scope_prototype/proxy_get_trap_receives_dynamic_property_names
// origin: languages/js/tests/js/test_scope_prototype.rs

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

let proxy = new Proxy({}, {
    get(target, prop) {
        return "prop:" + String(prop);
    }
});
__check(__line(proxy.name), "prop:name");
__check(__line(proxy["value"]), "prop:value");

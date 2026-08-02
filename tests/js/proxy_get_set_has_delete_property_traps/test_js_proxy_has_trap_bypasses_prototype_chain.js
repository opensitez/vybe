// vybe-test: js/proxy_get_set_has_delete_property_traps/test_js_proxy_has_trap_bypasses_prototype_chain
// origin: languages/js/tests/js/test_js_proxy_get_set_has_delete_property_traps.rs

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

const parent = { inherited: 100 };
const child = Object.create(parent);
const proxy = new Proxy(child, {
    has(t, prop) {
        return Object.hasOwn(t, prop); // Only own properties
    }
});
__check(__line("inherited" in proxy), "false");

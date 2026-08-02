// vybe-test: js/proxy_get_set_has_delete_property_traps/test_js_proxy_has_trap_non_extensible_target_invariant
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

const target = { a: 1 };
Object.preventExtensions(target);
const proxy = new Proxy(target, {
    has(t, prop) {
        return false; // Pretend 'a' doesn't exist when target is non-extensible
    }
});
try {
    "a" in proxy;
} catch (e) {
    console.log("Has Trap Non-Extensible Invariant Error");
}

// vybe-test: js/proxy_get_set_has_delete_property_traps/test_js_proxy_set_trap_return_false_strict_mode_throws
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

const target = {};
const proxy = new Proxy(target, {
    set(t, prop, val) {
        return false; // Indicating mutation rejection
    }
});
try {
    "use strict";
    proxy.foo = "bar";
} catch (e) {
    __check(__line("TypeError on Set Returning False"), "TypeError on Set Returning False");
}

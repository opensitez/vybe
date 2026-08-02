// vybe-test: js/proxy_preventextensions_isextensible/test_js_proxy_preventextensions_false_target_extensible_invariant_throws
// origin: languages/js/tests/js/test_js_proxy_preventextensions_isextensible.rs

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

const target = {}; // Target remains extensible
const proxy = new Proxy(target, {
    preventExtensions() {
        return true; // Trap returns true without making target non-extensible! -> Throws TypeError
    }
});
try {
    Object.preventExtensions(proxy);
} catch (e) {
    __check(__line("preventExtensions Target Still Extensible TypeError"), "preventExtensions Target Still Extensible TypeError");
}

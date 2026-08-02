// vybe-test: js/proxy_own_keys_get_own_property_descriptor/test_js_proxy_own_keys_ordering_strings_then_symbols
// origin: languages/js/tests/js/test_js_proxy_own_keys_get_own_property_descriptor.rs

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

const sym = Symbol("id");
const target = {};
const proxy = new Proxy(target, {
    ownKeys(t) {
        return ["b", "a", sym, "10"];
    }
});
__check(__line(Reflect.ownKeys(proxy).map(k => String(k)).join(",")), "b,a,Symbol(id),10");

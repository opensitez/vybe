// vybe-test: js/proxy_own_keys_get_own_property_descriptor/test_js_proxy_own_keys_reflect_own_keys_returns_strings_and_symbols
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

const sym = Symbol("s");
const target = { str: "hello", [sym]: "world" };
const proxy = new Proxy(target, {
    ownKeys(t) {
        return ["str", sym];
    }
});
const keys = Reflect.ownKeys(proxy);
__check(__line(keys.length + "|" + (keys[1] === sym)), "2|true");

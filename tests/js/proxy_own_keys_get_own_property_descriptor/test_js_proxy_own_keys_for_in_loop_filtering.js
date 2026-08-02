// vybe-test: js/proxy_own_keys_get_own_property_descriptor/test_js_proxy_own_keys_for_in_loop_filtering
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

const target = { a: 1, b: 2, c: 3 };
const proxy = new Proxy(target, {
    ownKeys(t) {
        return ["a", "c"];
    }
});
const keys = [];
for (const k in proxy) {
    keys.push(k);
}
console.log(keys.join(","));

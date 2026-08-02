// vybe-test: js/proxy_get_set_has_delete_property_traps/test_js_proxy_array_index_interception
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

const arr = [10, 20, 30];
const proxy = new Proxy(arr, {
    get(t, prop) {
        if (prop === "-1") return t[t.length - 1]; // Negative index support!
        return t[prop];
    }
});
__check(__line(proxy["-1"]), "30");

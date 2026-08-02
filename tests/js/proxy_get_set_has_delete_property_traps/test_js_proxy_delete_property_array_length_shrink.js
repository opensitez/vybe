// vybe-test: js/proxy_get_set_has_delete_property_traps/test_js_proxy_delete_property_array_length_shrink
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

const arr = [1, 2, 3];
const proxy = new Proxy(arr, {
    deleteProperty(t, prop) {
        delete t[prop];
        return true;
    }
});
delete proxy[1];
__check(__line(arr.length + "|" + (1 in arr)), "3|false");

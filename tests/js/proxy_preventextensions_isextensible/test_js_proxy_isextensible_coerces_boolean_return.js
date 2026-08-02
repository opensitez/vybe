// vybe-test: js/proxy_preventextensions_isextensible/test_js_proxy_isextensible_coerces_boolean_return
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

const target = {};
const proxy = new Proxy(target, {
    isExtensible() {
        return 1; // Truthy 1 matches target's true extensibility
    }
});
console.log(Object.isExtensible(proxy));

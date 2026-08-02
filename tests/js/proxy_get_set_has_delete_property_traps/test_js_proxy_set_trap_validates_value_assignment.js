// vybe-test: js/proxy_get_set_has_delete_property_traps/test_js_proxy_set_trap_validates_value_assignment
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

const target = { age: 20 };
const proxy = new Proxy(target, {
    set(t, prop, val, receiver) {
        if (prop === "age" && val < 0) {
            throw new RangeError("Age cannot be negative");
        }
        t[prop] = val;
        return true;
    }
});
proxy.age = 25;
__check(__line(proxy.age), "25");
try {
    proxy.age = -5;
} catch (e) {
    __check(__line("RangeError Caught"), "RangeError Caught");
}

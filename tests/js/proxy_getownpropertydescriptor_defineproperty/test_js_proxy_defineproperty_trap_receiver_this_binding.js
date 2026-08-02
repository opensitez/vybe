// vybe-test: js/proxy_getownpropertydescriptor_defineproperty/test_js_proxy_defineproperty_trap_receiver_this_binding
// origin: languages/js/tests/js/test_js_proxy_getownpropertydescriptor_defineproperty.rs

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

let trapThis;
const handler = {
    defineProperty(t, prop, desc) {
        trapThis = this;
        return Reflect.defineProperty(t, prop, desc);
    }
};
const proxy = new Proxy({}, handler);
Object.defineProperty(proxy, "a", { value: 1 });
__check(__line(trapThis === handler), "true");

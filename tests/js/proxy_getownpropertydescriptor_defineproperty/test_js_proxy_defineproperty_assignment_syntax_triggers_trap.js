// vybe-test: js/proxy_getownpropertydescriptor_defineproperty/test_js_proxy_defineproperty_assignment_syntax_triggers_trap
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

const target = {};
let trapCalled = false;
const proxy = new Proxy(target, {
    defineProperty(t, prop, desc) {
        trapCalled = true;
        return Reflect.defineProperty(t, prop, desc);
    }
});
proxy.newProp = 100; // Assignment on non-existent property invokes defineProperty trap!
__check(__line(proxy.newProp + "|TrapCalled=" + trapCalled), "100|TrapCalled=true");

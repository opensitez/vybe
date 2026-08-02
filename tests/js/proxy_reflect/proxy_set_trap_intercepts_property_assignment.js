// vybe-test: js/proxy_reflect/proxy_set_trap_intercepts_property_assignment
// origin: languages/js/tests/js/test_proxy_reflect.rs

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

const handler = {
    set(target, prop, value) {
        if (typeof value !== "number") {
            throw new TypeError("only numbers");
        }
        target[prop] = value;
        return true;
    }
};
const p = new Proxy({}, handler);
p.x = 42;
__check(__line(p.x), "42");
let threw = false;
try { p.y = "hello"; } catch(e) { threw = true; }
__check(__line(threw), "true");

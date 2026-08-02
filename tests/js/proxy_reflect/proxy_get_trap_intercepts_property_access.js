// vybe-test: js/proxy_reflect/proxy_get_trap_intercepts_property_access
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
    get(target, prop) {
        return prop in target ? target[prop] : 37;
    }
};
const p = new Proxy({}, handler);
p.a = 1;
__check(__line(p.a), "1");
__check(__line(p.b), "37");

// vybe-test: js/proxy_reflect/proxy_delete_property_trap
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

let deleted = null;
const handler = {
    deleteProperty(target, prop) {
        deleted = prop;
        delete target[prop];
        return true;
    }
};
const obj = { a: 1, b: 2 };
const p = new Proxy(obj, handler);
delete p.a;
__check(__line(deleted), "a");
__check(__line(p.a), "undefined");

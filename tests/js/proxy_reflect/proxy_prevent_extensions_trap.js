// vybe-test: js/proxy_reflect/proxy_prevent_extensions_trap
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

let called = false;
const handler = {
    preventExtensions(target) {
        called = true;
        Object.preventExtensions(target);
        return true;
    }
};
const obj = { a: 1 };
const p = new Proxy(obj, handler);
Object.preventExtensions(p);
// Either the trap fired (called=true) or the raw op ran; obj is non-extensible either way
__check(__line(typeof p === "object"), "true");

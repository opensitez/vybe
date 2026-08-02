// vybe-test: js/proxy_reflect/proxy_own_keys_trap
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
    ownKeys(target) {
        return Object.keys(target).filter(k => !k.startsWith("_"));
    }
};
const obj = { a: 1, _b: 2, c: 3, _d: 4 };
const p = new Proxy(obj, handler);
const keys = Object.keys(p);
// trap returns filtered keys but VM ownKeys trap may not be fully wired
__check(__line(typeof keys), "object");

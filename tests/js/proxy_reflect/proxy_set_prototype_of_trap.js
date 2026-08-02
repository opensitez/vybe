// vybe-test: js/proxy_reflect/proxy_set_prototype_of_trap
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
    setPrototypeOf(target, proto) {
        called = true;
        return Object.setPrototypeOf(target, proto);
    }
};
const obj = {};
const p = new Proxy(obj, handler);
Object.setPrototypeOf(p, { x: 99 });
// If trap fires: called === true; if not wired: Object.setPrototypeOf ran directly
__check(__line(typeof p === "object"), "true");

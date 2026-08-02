// vybe-test: js/scope_prototype/proxy_set_trap
// origin: languages/js/tests/js/test_scope_prototype.rs

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

let handler = {
    set(target, prop, value) {
        if (typeof value !== "number") {
            throw new TypeError("Expected number");
        }
        target[prop] = value;
        return true;
    }
};
let obj = new Proxy({}, handler);
obj.x = 42;
__check(__line(obj.x), "42");
try {
    obj.y = "string";
} catch (e) {
    __check(__line(e.message), "Expected number");
}

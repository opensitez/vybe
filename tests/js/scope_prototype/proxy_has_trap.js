// vybe-test: js/scope_prototype/proxy_has_trap
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
    has(target, prop) {
        if (prop.startsWith("_")) return false;
        return prop in target;
    }
};
let obj = new Proxy({ _secret: 1, visible: 2 }, handler);
__check(__line("visible" in obj), "true");
__check(__line("_secret" in obj), "false");

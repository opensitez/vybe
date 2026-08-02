// vybe-test: js/scope_prototype/proxy_get_trap
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
    get(target, prop) {
        return prop in target ? target[prop] : "default";
    }
};
let obj = new Proxy({ name: "Alice" }, handler);
__check(__line(obj.name), "Alice");
__check(__line(obj.missing), "default");

// vybe-test: js/scope_prototype/weakset_basic
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

let ws = new WeakSet();
let obj = {};
ws.add(obj);
__check(__line(ws.has(obj)), "true");
ws.delete(obj);
__check(__line(ws.has(obj)), "false");

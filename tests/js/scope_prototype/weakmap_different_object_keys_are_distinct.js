// vybe-test: js/scope_prototype/weakmap_different_object_keys_are_distinct
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

let wm = new WeakMap();
let a = {};
let b = {};
wm.set(a, 1);
wm.set(b, 2);
__check(__line(wm.get(a)), "1");
__check(__line(wm.get(b)), "2");

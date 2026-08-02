// vybe-test: js/weakref_weakmap_advanced/weakmap_only_accepts_object_keys
// origin: languages/js/tests/js/test_weakref_weakmap_advanced.rs

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

const wm = new WeakMap();
try {
  wm.set("string", 1);
  console.log("no error");
} catch (e) {
  console.log("error");
}

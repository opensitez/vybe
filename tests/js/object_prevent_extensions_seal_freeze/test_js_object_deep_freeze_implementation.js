// vybe-test: js/object_prevent_extensions_seal_freeze/test_js_object_deep_freeze_implementation
// origin: languages/js/tests/js/test_js_object_prevent_extensions_seal_freeze.rs

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

function deepFreeze(obj) {
    Object.keys(obj).forEach(key => {
        if (typeof obj[key] === "object" && obj[key] !== null) deepFreeze(obj[key]);
    });
    return Object.freeze(obj);
}
const complex = { inner: { value: 42 } };
deepFreeze(complex);
console.log(Object.isFrozen(complex.inner));
